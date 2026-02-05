# -*- coding: utf-8 -*-

# HEADER
__title__ = "Exportar NV1"
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
NOMEPECA = "Modelo"
MARCA = "Marca"
CODCONTROLE = ""
DESENHO = ""
TIPOPRODUTO = "03. PRODUTO"
GRUPO = "04. GRUPO"
SECAO = "05. SEÇÃO"
INFOADICIONAL = "09. INFO ADICIONAL"
COMPRIMENTO = "08. COMPRIMENTO"
ALTURA = "07. ALTURA"
LARGURA = "06. LARGURA"
VOLUMEUNITARIO = "Volume"
PESO = "Peso"
AREA = ""
CLASSECONCRETO = "12. FCK"
ACABAMENTO = ""
COBRIMENTO = "13. COBRIMENTO"
OBS = ""
TABELAACO = ""
COMPLEMENTOS = ""
OBRA = "Obra da Serpa"
OBRAPARAM = ""
NAME = "Serpa"
PROJETISTA = "Projetista"

# FUNCTIONS
def reject_invalid(list_of_elements):
    accepted_elements = []
    rejected_count = 0
    # >>> CATEGORIAS DE INTERESSE <<<
    categorias_interesse = [
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Assemblies,
        # BuiltInCategory.OST_GenericModel,
    ]
    categorias_ids = [int(cat) for cat in categorias_interesse]

    for element in list_of_elements:
        if not element.Category:
            rejected_count += 1
            continue
        if element.Category.Id.IntegerValue in categorias_ids:
            accepted_elements.append(element)
        else:
            rejected_count += 1
    if not accepted_elements:
        print(
            "Nenhum elemento selecionado faz parte das categorias de interesse da exportação. Faça uma nova seleção e "
            "tente novamente."
        )
        sys.exit(0)
    if rejected_count > 0:
        print(
            "Elementos que não pertencem às categorias de interesse da exportação foram removidos. A execução do "
            "código seguirá com os elementos válidos."
        )
    return accepted_elements


def group_elements(elements):
    grupos = {}
    for element in elements:
        nomepeca = get_nome_peca(element)
        tipoproduto = parameter_get(element, TIPOPRODUTO)
        grupo = parameter_get(element, GRUPO)
        secao = parameter_get(element, SECAO)
        infoadicional = parameter_get(element, INFOADICIONAL)
        comprimento = parameter_get(element, COMPRIMENTO)
        altura = parameter_get(element, ALTURA)
        largura = parameter_get(element, LARGURA)
        volumeunitario = parameter_get(element, VOLUMEUNITARIO)
        peso = parameter_get(element, PESO)
        area = parameter_get(element, AREA)
        classeconcreto = parameter_get(element, CLASSECONCRETO)
        acabamento = parameter_get(element, ACABAMENTO)
        cobrimento = parameter_get(element, COBRIMENTO)
        obs = parameter_get(element, OBS)
        tabelaaco = parameter_get(element, TABELAACO)
        complementos = parameter_get(element, COMPLEMENTOS)
        chave = (
            nomepeca,
            tipoproduto,
            grupo,
            secao,
            infoadicional,
            comprimento,
            altura,
            largura,
            volumeunitario,
            peso,
            area,
            classeconcreto,
            acabamento,
            cobrimento,
            obs,
            tabelaaco,
            complementos,
        )
        if chave not in grupos:
            grupos[chave] = {
                "elemento_base": element,
                "quantidade": 1,
                "ids": [element.UniqueId]
            }
        else:
            grupos[chave]["quantidade"] += 1
            grupos[chave]["ids"].append(element.UniqueId)
    return grupos


def get_main_element(element):
    categorias_principais = [
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Floors,
    ]
    if not isinstance(element, AssemblyInstance):
        return element
    for member_id in element.GetMemberIds():
        member = doc.GetElement(member_id)
        if not member or not member.Category:
            continue
        if member.Category.Id.IntegerValue in [int(c) for c in categorias_principais]:
            return member
    print(
        "A montagem '{}' não possui uma peça principal válida para exportação."
        .format(element.Id)
    )
    return None


def get_nome_peca(element):
    modelo = parameter_get(element, NOMEPECA)
    marca = parameter_get(element, MARCA)

    if modelo and marca:
        return "{}{}".format(modelo, marca)
    elif modelo:
        return modelo
    else:
        return str(element.Id)


def clean_xml_text(value, parameter_name, element):
    if value is None:
        return ""

    original_value = value
    sanitized_value = value

    caracteres_invalidos = ['&', '<', '>', '"', "'"]

    for c in caracteres_invalidos:
        if c in sanitized_value:
            sanitized_value = sanitized_value.replace(c, "")
            nome_peca = get_nome_peca(element)
            print(
                "O caractere '{}' do texto '{}' do parâmetro '{}' da peça '{}' "
                "foi removido do xml para preservar a validade do arquivo."
                .format(c, original_value, parameter_name, nome_peca)
            )

    return sanitized_value


def parameter_get(element, parameter_name):
    if parameter_name == "":
        # print("Parâmetro ignorado.")
        return ""
    param_instance = element.LookupParameter(parameter_name)
    if param_instance:
        if parameter_name == VOLUMEUNITARIO:
            try:
                volume_ft3 = param_instance.AsDouble()
                volume_m3 = volume_ft3 * 0.028316846592
                # print(
                #     "Parâmetro de instância '{}' encontrado. Valor: {}"
                #     .format(parameter_name, volume_m3)
                # )
                return "{:.3f}".format(volume_m3)
            except:
                return ""
        if parameter_name == PESO:
            try:
                valor = param_instance.AsDouble()
                # print(
                #     "Parâmetro de instância '{}' encontrado. Valor: {}"
                #     .format(parameter_name, valor)
                # )
                return "{:.3f}".format(valor)
            except:
                value = param_instance.AsValueString()
                if not value:
                    return ""
                # print(
                #     "Parâmetro de instância '{}' encontrado. Valor: {}"
                #     .format(parameter_name, value)
                # )
                return clean_xml_text(value.replace(",", "."), parameter_name, element)
        value = param_instance.AsValueString()
        if not value:
            return ""
        # print(
        #     "Parâmetro de instância '{}' encontrado. Valor: {}"
        #     .format(parameter_name, value)
        # )
        return clean_xml_text(value.replace(",", "."), parameter_name, element)
    type_id = element.GetTypeId()
    if type_id and type_id != ElementId.InvalidElementId:
        element_type = doc.GetElement(type_id)
        if element_type:
            param_type = element_type.LookupParameter(parameter_name)
            if param_type:
                if parameter_name == VOLUMEUNITARIO:
                    try:
                        volume_ft3 = param_type.AsDouble()
                        volume_m3 = volume_ft3 * 0.028316846592
                        # print(
                        #     "Parâmetro de tipo '{}' encontrado. Valor: {}"
                        #     .format(parameter_name, volume_m3)
                        # )
                        return "{:.3f}".format(volume_m3)
                    except:
                        return ""

                if parameter_name == PESO:
                    try:
                        valor = param_type.AsDouble()
                        # print(
                        #     "Parâmetro de tipo '{}' encontrado. Valor: {}"
                        #     .format(parameter_name, valor)
                        # )
                        return "{:.3f}".format(valor)
                    except:
                        value = param_type.AsValueString()
                        if not value:
                            return ""
                        # print(
                        #     "Parâmetro de tipo '{}' encontrado. Valor: {}"
                        #     .format(parameter_name, value)
                        # )
                        return clean_xml_text(value.replace(",", "."), parameter_name, element)

                value = param_type.AsValueString()
                if not value:
                    return ""
                # print(
                #     "Parâmetro de tipo '{}' encontrado. Valor: {}"
                #     .format(parameter_name, value)
                # )
                return clean_xml_text(value.replace(",", "."), parameter_name, element)
    # print(
    #     "Parâmetro '{}' não foi encontrado."
    #     .format(parameter_name)
    # )
    return ""


def filter_elements(list_of_elements):
    accepted_output = []
    parametros_obrigatorios = [
        NOMEPECA,
        TIPOPRODUTO,
        GRUPO,
        SECAO,
        INFOADICIONAL,
        COMPRIMENTO,
        VOLUMEUNITARIO
    ]
    for element in list_of_elements:
        elemento_valido = True
        nome_para_print = get_nome_peca(element)
        for parametro in parametros_obrigatorios:
            # Busca instância ou tipo
            param_instancia = element.LookupParameter(parametro)
            param_tipo = None
            if not param_instancia:
                type_id = element.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    element_type = doc.GetElement(type_id)
                    if element_type:
                        param_tipo = element_type.LookupParameter(parametro)

            param = param_instancia or param_tipo
            if not param:
                print(
                    "O elemento '{}' foi removido porque não possui o parâmetro "
                    "obrigatório '{}' definido."
                    .format(nome_para_print, parametro)
                )
                elemento_valido = False
                break
            valor = param.AsValueString()
            if valor is None or valor == "":
                print(
                    "O elemento '{}' foi removido porque está com o parâmetro "
                    "obrigatório '{}' vazio."
                    .format(nome_para_print, parametro)
                )
                elemento_valido = False
                break
        if elemento_valido:
            accepted_output.append(element)
    return accepted_output


# def group_elements(list_of_elements):
#     for element in list_of_elements:


def xml_unit_build(selected_element, quantidade, ids):
    nomepeca = get_nome_peca(selected_element)
    codcontrole = parameter_get(selected_element, CODCONTROLE)
    desenho = parameter_get(selected_element, DESENHO)
    tipoproduto = parameter_get(selected_element, TIPOPRODUTO)
    grupo = parameter_get(selected_element, GRUPO)
    secao = parameter_get(selected_element, SECAO)
    infoadicional = parameter_get(selected_element, INFOADICIONAL)
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
    xml_quantidade = "\t\t<QUANTIDADE>{}</QUANTIDADE>\n".format(quantidade)
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
    lista_ids = ""
    for uid in ids:
        lista_ids += "\t\t\t<ID>{}</ID>\n".format(uid)
    xml_listaid = "\t\t<LISTAID>\n" + lista_ids + "\t\t</LISTAID>\n"
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

    main_element = get_main_element(elem)

    if main_element:
        selected_elements.append(main_element)

if not selected_elements:
    print("Nenhum elemento foi selecionado no modelo. Favor selecionar e tentar novamente.")
    sys.exit(1)
else:
    pass

valid_elements = reject_invalid(selected_elements)
filtered_elements = filter_elements(valid_elements)
# group_elements(selected_elements)
if not filtered_elements:
    print(
        "\nNenhuma peça válida foi encontrada.\n"
        "O arquivo XML não foi gerado.\n"
        "Verifique se os parâmetros obrigatórios estão preenchidos corretamente."
    )
    sys.exit(0)

# 2. Define the output directory path:
#directory_path = r"C:\Users\Viegas\Desktop\Profissional\Freelance\3. Serpa\Output"
#directory_path = r"C:\Users\Viegas\Desktop\Profissional\Freelance\4. Concrete Show\Output"
rvt_path = doc.PathName

if not rvt_path:
    print(
        "\nO arquivo do Revit ainda não foi salvo.\n"
        "Salve o arquivo antes de exportar o XML."
    )
    sys.exit(1)

directory_path = os.path.dirname(rvt_path)

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
    grupos = group_elements(filtered_elements)

    for grupo in grupos.values():
        xml_unit = xml_unit_build(
            grupo["elemento_base"],
            grupo["quantidade"],
            grupo["ids"]
        )
        xml_file.write(xml_unit)
    xml_file.write(xml_detalhamento_close)

if output_space == 1:
    print('\nElementos válidos exportados com sucesso para o documento "{}" dentro do diretório "{}".'
          .format(output_string, directory_path))
else:
    print('Elementos válidos exportados com sucesso para o documento "{}" dentro do diretório "{}".'
          .format(output_string, directory_path))
