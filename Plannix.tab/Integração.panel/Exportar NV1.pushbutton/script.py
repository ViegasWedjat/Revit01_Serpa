# -*- coding: utf-8 -*-

# HEADER
__title__ = "Exportar XML"
__author__ = "Daniel Viegas"
__doc__ = """Selecione os elementos de interesse do modelo tridimensional do Revit e exporte os seus dados 
parametrizados para um arquivo XML para ser integrado no software Plannix."""

# IMPORTS
import clr

clr.AddReference("System")
from System.Collections.Generic import List
import os
import sys
import math
import datetime
import time
from Autodesk.Revit.DB import *
from pyrevit import revit, forms
import xml.etree.ElementTree as ET

# HARD VARIABLES
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
output_space = 0

# SOFT VARIABLES
OBRA = "Obra da Serpa"
NAME = "Serpa"
PROJETISTA = "Equipe da Serpa"
NOMEPECA = "01. Nome da Peça"
CODCONTROLE = "02. Código de Controle"
DESENHO = "03. Desenho da Peça"
TIPOPRODUTO = "04. Produto"
GRUPO = "05. Grupo"
SECAO = "06. Seção"
INFOADICIONAL = "07. Info. Adicional"
COMPRIMENTO = "08. Comprimento"
ALTURA = "09. Altura"
LARGURA = "10. Largura"
VOLUMEUNITARIO = "11. Volume"
PESO = "12. Peso"
AREA = "13. Área"
CLASSECONCRETO = "14. Classe"
ACABAMENTO = "15. Acabamento"
COBRIMENTO = "16. Cobrimento"
OBS = "17. Observação"
TABELAACO = "18. Tabela de Aço"
COMPLEMENTOS = "19. Complementos"


# FUNCTIONS
def reject_invalid(list_of_elements):
    rejected_elements = []
    accepted_elements = []
    for element in list_of_elements:
        parameters = element.Parameters
        name_check = NOMEPECA
        count = 1
        for parameter in parameters:
            if parameter.Definition.Name == name_check:
                count -= 1
        if count == 1:
            rejected_elements.append(element)
        else:
            accepted_elements.append(element)
    return accepted_elements


def check_parameter(check):
    if check.AsValueString() is None:
        return 0
    elif check.AsValueString() == "":
        return 0
    else:
        return 1


def filter_elements(list_of_elements):
    global output_space
    rejected_output = []
    accepted_output = []
    for element in list_of_elements:
        check1 = element.LookupParameter(NOMEPECA)
        check2 = element.LookupParameter(TIPOPRODUTO)
        check3 = element.LookupParameter(GRUPO)
        check4 = element.LookupParameter(SECAO)
        check5 = element.LookupParameter(INFOADICIONAL)
        check6 = element.LookupParameter(COMPRIMENTO)
        check7 = element.LookupParameter(VOLUMEUNITARIO)
        if check_parameter(check1) == 0:
            rejected_output.append(element)
        elif check_parameter(check2) == 0:
            rejected_output.append(element)
        elif check_parameter(check3) == 0:
            rejected_output.append(element)
        elif check_parameter(check4) == 0:
            rejected_output.append(element)
        elif check_parameter(check5) == 0:
            rejected_output.append(element)
        elif check_parameter(check6) == 0:
            rejected_output.append(element)
        elif check_parameter(check7) == 0:
            rejected_output.append(element)
        else:
            accepted_output.append(element)
    for rejected_element in rejected_output:
        id_string = str(rejected_element.Id)
        print("O elemento {} não tem todos os parâmetros obrigatórios informados. Favor conferir.".format(id_string))
    if rejected_output:
        print("\nOs parâmetros obrigatórios são:\n{};\n{};\n{};\n{};\n{};\n{};\n{}.\n\n"
              .format(NOMEPECA, TIPOPRODUTO, GRUPO, SECAO, INFOADICIONAL, COMPRIMENTO, VOLUMEUNITARIO))
        output_space = 1

    return accepted_output


# def group_elements(list_of_elements):
#     for element in list_of_elements:


def parameter_get(element, parameter_name):
    parameter = element.LookupParameter(parameter_name)
    if parameter:
        return parameter.AsValueString() or ""


def xml_unit_build(selected_element):
    nomepeca = parameter_get(selected_element, NOMEPECA)
    codcontrole = parameter_get(selected_element, CODCONTROLE)
    desenho = parameter_get(selected_element, DESENHO)
    tipoproduto = parameter_get(selected_element, TIPOPRODUTO)
    grupo = parameter_get(selected_element, GRUPO)
    secao = parameter_get(selected_element, SECAO)
    infoadicional = parameter_get(selected_element, INFOADICIONAL)
    quantidade = r"1"
    comprimento = parameter_get(selected_element, COMPRIMENTO)
    altura = parameter_get(selected_element, ALTURA)
    largura = parameter_get(selected_element, LARGURA)
    volumeunitario = parameter_get(selected_element, VOLUMEUNITARIO)
    peso = parameter_get(selected_element, PESO)
    area = parameter_get(selected_element, AREA)
    classeconcreto = parameter_get(selected_element, CLASSECONCRETO)
    acabamento = parameter_get(selected_element, ACABAMENTO)
    cobrimento = parameter_get(selected_element, COBRIMENTO)
    obs = parameter_get(selected_element, OBS)
    listaid = "\t\t\t\t\t<ID>" + selected_element.UniqueId + "</ID>\n"
    tabelaaco = parameter_get(selected_element, TABELAACO)
    complementos = parameter_get(selected_element, COMPLEMENTOS)
    xml_peca_open = "\t<PECA>\n"
    xml_peca_close = "\t</PECA>\n"
    xml_nomepeca = "\t\t<NOMEPECA>" + nomepeca + "</NOMEPECA>\n"
    xml_codcontrole = "\t\t<CODCONTROLE>" + codcontrole + "</CODCONTROLE>\n"
    xml_desenho = "\t\t<DESENHO>" + desenho + "</DESENHO>\n"
    xml_tipoproduto = "\t\t<TIPOPRODUTO>" + tipoproduto + "</TIPOPRODUTO>\n"
    xml_grupo = "\t\t<GRUPO>" + grupo + "</GRUPO>\n"
    xml_secao = "\t\t<SECAO>" + secao + "</SECAO>\n"
    xml_infoadicional = "\t\t<INFOADICIONAL>" + infoadicional + "</INFOADICIONAL>\n"
    xml_quantidade = "\t\t<QUANTIDADE>" + quantidade + "</QUANTIDADE>\n"
    xml_comprimento = "\t\t<COMPRIMENTO>" + comprimento + "</COMPRIMENTO>\n"
    xml_altura = "\t\t<ALTURA>" + altura + "</ALTURA>\n"
    xml_largura = "\t\t<LARGURA>" + largura + "</LARGURA>\n"
    xml_volumeunitario = "\t\t<VOLUMEUNITARIO>" + volumeunitario + "</VOLUMEUNITARIO>\n"
    xml_peso = "\t\t<PESO>" + peso + "</PESO>\n"
    xml_area = "\t\t<AREA>" + area + "</AREA>\n"
    xml_classeconcreto = "\t\t<CLASSECONCRETO>" + classeconcreto + "</CLASSECONCRETO>\n"
    xml_acabamento = "\t\t<ACABAMENTO>" + acabamento + "</ACABAMENTO>\n"
    xml_cobrimento = "\t\t<COBRIMENTO>" + cobrimento + "</COBRIMENTO>\n"
    xml_obs = "\t\t<OBS>" + obs + "</OBS>\n"
    xml_listaid = "\t\t<LISTAID>\n" + listaid + "\t\t</LISTAID>\n"
    xml_tabelaaco = "\t\t<TABELAACO>\n" + tabelaaco + "\t\t</TABELAACO>\n"
    xml_complementos = "\t\t<COMPLEMENTOS>\n" + complementos + "\t\t</COMPLEMENTOS>\n"
    element_unit_string = (xml_peca_open + xml_nomepeca + xml_codcontrole + xml_desenho + xml_tipoproduto + xml_grupo
                           + xml_secao + xml_infoadicional + xml_quantidade + xml_comprimento + xml_altura +
                           xml_largura + xml_volumeunitario + xml_peso + xml_area + xml_classeconcreto +
                           xml_acabamento + xml_cobrimento + xml_obs + xml_listaid + xml_tabelaaco + xml_complementos
                           + xml_peca_close)
    return element_unit_string


# MAIN CODE

# 1. Select the elements and part them into unique elements or grouped (repeated) elements:

selected_elements = []

for elem_id in uidoc.Selection.GetElementIds():
    elem = uidoc.Document.GetElement(elem_id)
    selected_elements.append(elem)

if not selected_elements:
    print("Nenhum elemento foi selecionado no modelo. Favor selecionar e tentar novamente.")
    sys.exit(1)
else:
    pass

valid_elements = reject_invalid(selected_elements)
filtered_elements = filter_elements(valid_elements)
# group_elements(selected_elements)

# 2. Define the output directory path:
directory_path = r"C:\Users\Viegas\Desktop\Profissional\Freelance\3. Serpa\Output"

# 3. Define the output file:
current_datetime = datetime.datetime.now()
datetime_string = current_datetime.strftime("[%Y-%m-%d][%Hh%Mm]")
output_string = "Export" + datetime_string + ".xml"
xml_file_path = os.path.join(directory_path, output_string)

# 4. Build the xml basic structure:
xml_header = '<?xml version="1.0" encoding="ISO-8859-1" ?>\n'
xml_detalhamento_open = ('<DETALHAMENTOPLANNIX obra="' + OBRA + '" name="' + NAME + '" projetista="' + PROJETISTA +
                         '">\n')
xml_detalhamento_close = '</DETALHAMENTOPLANNIX>'

# 5. Export the structured xml file:
with open(xml_file_path, "w") as xml_file:
    xml_file.write(xml_header)
    xml_file.write(xml_detalhamento_open)
    for element_unit in filtered_elements:
        xml_unit = xml_unit_build(element_unit)
        xml_file.write(xml_unit)
    xml_file.write(xml_detalhamento_close)

if output_space == 1:
    print('\nElementos válidos exportados com sucesso para o documento "{}" dentro do diretório "{}".'
          .format(output_string, directory_path))
else:
    print('Elementos válidos exportados com sucesso para o documento "{}" dentro do diretório "{}".'
          .format(output_string, directory_path))
