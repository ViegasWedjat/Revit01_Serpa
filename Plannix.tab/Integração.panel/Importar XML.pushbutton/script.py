# -*- coding: utf-8 -*-

# HEADER
__title__ = "Importar XML"
__author__ = "Daniel Viegas"
__doc__ = """Selecione um arquivo XML exportado do software Plannix e importe os dados de código de controle, status e data do status para os elementos correspondentes no modelo do Revit."""

# IMPORTS
import os
import sys
import xml.etree.ElementTree as ET
import clr
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
from System.Collections.Generic import List
from System.Windows.Forms import OpenFileDialog, DialogResult
from Autodesk.Revit.DB import *
from pyrevit import script

# HARD VARIABLES
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# PARAM NAMES
PARAM_CODCONTROLE = "16. CÓDIGO DE CONTROLE"
PARAM_STATUS      = "18. STATUS DA PEÇA"
PARAM_DATA        = "19. DATA DO STATUS"

# TRADUÇÃO STATUS PLANNIX → REVIT
STATUS_MAP = {
    "PROJETADA"       : "Projetada",
    "PROGRAMADA"      : "Programada",
    "CORTE E DOBRA"   : "Corte e Dobra Realizado",
    "ARMAÇÃO"         : "Armação Realizada",
    "FORMA"           : "Forma Realizada",
    "FORMA E ARMAÇÃO" : "Forma com Armação Realizada",
    "CONCRETAGEM"     : "Concretagem Realizada",
    "PREPARAÇÃO"      : "Preparação Realizada",
    "CORTE"           : "Corte Realizado",
    "PRÉ MONTAGEM"    : "Pré-montagem Realizada",
    "MONTAGEM"        : "Montagem Realizada (Met.)",
    "SOLDA"           : "Solda Realizada",
    "ACABAMENTO"      : "Acabamento Realizado (Met.)",
    "JATEAMENTO"      : "Jateamento Realizado",
    "GALVANIZAÇÃO"    : "Galvanização Realizada",
    "PINTURA"         : "Pintura Realizada",
    "ACABADA"         : "Acabamento Realizado",
    "EXPEDIDA"        : "Expedida para a Obra",
    "DEVOLVIDA"       : "Devolvida pela Obra",
    "DESCARREGADA"    : "Descarregada na Obra",
    "MONTADA"         : "Montada na Obra",
}

# FUNCTIONS
def set_parameter(element, param_name, value, nomepeca, guid, label):
    param = element.LookupParameter(param_name)
    if not param:
        print(
            "O parâmetro para inserir a informação de {} não foi encontrado na peça '{}' de GUID {}.".format(
                label, nomepeca, guid
            )
        )
        return False
    try:
        param.Set(value)
        return True
    except Exception as e:
        print(
            "Erro ao definir '{}' na peça '{}' de GUID {}: {}".format(
                param_name, nomepeca, guid, str(e)
            )
        )
        return False


# MAIN CODE

# 1. Selecionar arquivo XML
rvt_path = doc.PathName
initial_dir = os.path.dirname(rvt_path) if rvt_path else os.path.expanduser("~")

dialog = OpenFileDialog()
dialog.Title        = "Selecione o arquivo XML do Plannix"
dialog.Filter       = "Arquivos XML (*.xml)|*.xml"
dialog.FilterIndex  = 1
dialog.Multiselect  = False
dialog.InitialDirectory = initial_dir

result = dialog.ShowDialog()
if result != DialogResult.OK or not dialog.FileName:
    print("Nenhum arquivo selecionado. Execução cancelada.")
    sys.exit(0)

xml_path = dialog.FileName

# 2. Parsear XML
try:
    tree = ET.parse(xml_path)
    root = tree.getroot()
except Exception as e:
    print("Erro ao ler o arquivo XML: {}".format(str(e)))
    sys.exit(1)

pecas = root.findall("PECA")
if not pecas:
    print("Nenhuma peça encontrada no XML.")
    sys.exit(0)

# 3. Processar peças dentro de transação
count_ok   = 0
count_skip = 0
count_err  = 0

try:
    with Transaction(doc, "Importar XML Plannix") as t:
        t.Start()
        for peca in pecas:

            nomepeca = (peca.findtext("NOMEPECA") or "").strip()
            guid     = (peca.findtext("ID")          or "").strip()
            cod      = (peca.findtext("CODCONTROLE") or "").strip()
            status_raw = (peca.findtext("STATUS") or "").strip()
            status     = STATUS_MAP.get(status_raw.upper(), status_raw)
            data     = (peca.findtext("DATA")        or "").strip()

            # ID vazio → pular silenciosamente
            if not guid:
                count_skip += 1
                continue

            # Buscar elemento pelo GUID
            element = doc.GetElement(guid)
            if not element:
                print(
                    "A peça '{}' de GUID {} não foi encontrada no modelo e será desconsiderada na execução.".format(
                        nomepeca, guid
                    )
                )
                count_err += 1
                continue

            # Checkout do elemento (modelos workshared)
            if doc.IsWorkshared:
                try:
                    checkout_ids = List[ElementId]([element.Id])
                    WorksharingUtils.CheckoutElements(doc, checkout_ids)
                except Exception as e:
                    error_msg = str(e).lower()
                    if "central" in error_msg or "network" in error_msg or "reached" in error_msg or "server" in error_msg:
                        print(
                            "O modelo central está inacessível na rede. "
                            "Não é possível editar elementos em modelos workshared sem conexão com o arquivo central. "
                            "Abra o arquivo destacado do central ou conecte-se ao servidor e tente novamente."
                        )
                        t.RollBack()
                        sys.exit(1)
                    print(
                        "Não foi possível fazer checkout da peça '{}' de GUID {}: {}".format(
                            nomepeca, guid, str(e)
                        )
                    )
                    count_err += 1
                    continue

            # Preencher parâmetros
            ok_cod    = set_parameter(element, PARAM_CODCONTROLE, cod,    nomepeca, guid, "código de controle")
            ok_status = set_parameter(element, PARAM_STATUS,      status, nomepeca, guid, "status da peça")
            ok_data   = set_parameter(element, PARAM_DATA,        data,   nomepeca, guid, "data do status")

            if ok_cod and ok_status and ok_data:
                count_ok += 1
            else:
                count_err += 1

        t.Commit()

except Exception as e:
    error_msg = str(e).lower()
    if "central" in error_msg or "workset" in error_msg or "inaccessible" in error_msg or "worksharing" in error_msg:
        print(
            "O documento atual não pode ser editado pois o modelo central está inacessível para o usuário atual. "
            "Verifique se o nome de usuário do Revit corresponde ao proprietário do modelo ou solicite acesso ao administrador."
        )
    else:
        print("Erro inesperado durante a importação: {}".format(str(e)))
    sys.exit(1)

# 4. Resumo
if count_ok > 0:
    print("Elementos válidos do modelo preenchidos com sucesso com os dados válidos do XML.")

if count_err > 0 or count_skip > 0:
    print(
        "\nResumo da importação:\n"
        "  Peças atualizadas com sucesso  : {}\n"
        "  Peças com erros/não encontradas: {}\n"
        "  Peças ignoradas (ID vazio)     : {}".format(count_ok, count_err, count_skip)
    )