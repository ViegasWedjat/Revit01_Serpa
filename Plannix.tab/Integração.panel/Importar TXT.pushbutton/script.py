# -*- coding: utf-8 -*-

__title__ = "Importar TXT"
__author__ = "Daniel Viegas"
__doc__ = """Selecione o documento de referência e importe os dados do software Plannix para o seu modelo
tridimensional do Revit."""

# if __name__ == '__main__':
#     print("Funcionalidade em desenvolvimento!")

# IMPORTS
import re
import clr
clr.AddReference("System")
from System.Collections.Generic import List
import os
import sys
from Autodesk.Revit.DB import *
from pyrevit import revit, forms

# HARD VARIABLES
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)

# SOFT VARIABLES
CODCONTROLE = "02. Código de Controle"
STATUSPECA = "20. Status da Peça"
DATASTATUS = "21. Data do Status"
status01 = "Projetada"
status02 = "Programada"
status03 = "Corte e Dobra Realizado"
status04 = "Armação Realizada"
status05 = "Forma Realizada"
status06 = "Forma com Armação Realizada"
status07 = "Concretagem Realizada"
status08 = "Acabamento Realizado"
status09 = "Expedida para a Obra"
status10 = "Devolvida pela Obra"
status11 = "Descarregada na Obra"
status12 = "Montada na Obra"

# FUNCTIONS
def remove_spaces(input_string):
    pattern = r'(\w+)\s*$'
    match = re.search(pattern, input_string)
    if match:
        alphanumeric_substr = match.group(1)
        last_index = input_string.rfind(alphanumeric_substr)
        return input_string[:last_index + len(alphanumeric_substr)]
    else:
        return input_string


def check_status(statuscode):
    global status01
    global status02
    global status03
    global status04
    global status05
    global status06
    global status07
    global status08
    global status09
    global status10
    global status11
    global status12
    if statuscode == "000001":
        return status01
    elif statuscode == "000002":
        return status02
    elif statuscode == "000003":
        return status03
    elif statuscode == "000004":
        return status04
    elif statuscode == "000005":
        return status05
    elif statuscode == "000006":
        return status06
    elif statuscode == "000007":
        return status07
    elif statuscode == "000008":
        return status08
    elif statuscode == "000009":
        return status09
    elif statuscode == "000010":
        return status10
    elif statuscode == "000011":
        return status11
    elif statuscode == "000012":
        return status12


def element_by_uniqueid(doc, uniqueid):
    global valid
    try:
        element = doc.GetElement(uniqueid)
        return element
    except (ValueError, Exception):
        valid -= 1
        return None


def set_parameter_value(element, parameter_name, value):
    parameter = element.LookupParameter(parameter_name)
    if parameter:
        parameter.Set(value)


# MAIN CODE

# 1. Define the file path:
file_path = r"C:\Users\Viegas\Desktop\Profissional\Freelance\3. Serpa\Input\PlannixExport.txt"

# 2. Open the file and read its content as lines:
with open(file_path, 'r') as file:
    file_lines = file.readlines()

# 3. Remove header and '\n' from the lines list:
file_noheader = file_lines[1:]
lines = []
for line in file_noheader:
    provisory = line.replace("\n", "")
    lines.append(provisory)

# 4. Split string into 4 pieces:
linestrings = []
for line in lines:
    valid = 1
    linelength = len(line)
    linestring = []
    if line[0].isalnum():
        if line[29].isspace():
            linename = line[:30]
            adjustedlinename = remove_spaces(linename)
            linestring.append(adjustedlinename)
        else:
            valid -= 1
    else:
        valid -= 1
    if line[30].isalnum():
        if line[59].isspace():
            linecode = line[30:60]
            adjustedlinecode = remove_spaces(linecode)
            linestring.append(adjustedlinecode)
        else:
            valid -= 1
    else:
        valid -= 1
    if line[60].isalnum():
        if line[(linelength - 16)].isalnum():
            lineguid = line[60:(linelength - 15)]
            adjustedlineguid = remove_spaces(lineguid)
            detected_element = element_by_uniqueid(doc, adjustedlineguid)
            linestring.append(detected_element)
        else:
            valid -= 1
    else:
        valid -= 1
    if line[(linelength - 15)].isalnum():
        if line[(linelength - 1)].isalnum():
            linestatus = line[(linelength - 15):(linelength - 9)]
            adjustedlinestatus = remove_spaces(linestatus)
            statusplannix = check_status(adjustedlinestatus)
            linestring.append(statusplannix)
            linedate = line[(linelength - 8):(linelength)]
            adjustedlinedate = remove_spaces(linedate)
            linestring.append(adjustedlinedate)
        else:
            valid -= 1
    else:
        valid -= 1
    if valid == 1:
        linestrings.append(linestring)

# 5. Start transaction, set the parameters values and commit the transaction:
t = Transaction(doc, __title__)

t.Start()

for line in linestrings:
    set_parameter_value(line[2], CODCONTROLE, line[1])
    set_parameter_value(line[2], STATUSPECA, line[3])
    set_parameter_value(line[2], DATASTATUS, line[4])

t.Commit()
