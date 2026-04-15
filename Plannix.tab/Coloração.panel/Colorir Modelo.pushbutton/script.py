# -*- coding: utf-8 -*-

# HEADER
__title__ = "Colorir Modelo"
__author__ = "Daniel Viegas"
__doc__ = """Aplica, remove cores e oculta/reexibe categorias nos elementos do modelo na vista ativa de acordo com o status da peça."""

# IMPORTS
import os
import io
import sys
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from pyrevit import forms, script

# HARD VARIABLES
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
view        = uidoc.ActiveView
output      = script.get_output()
PATH_SCRIPT = os.path.dirname(__file__)
xaml_path   = os.path.join(PATH_SCRIPT, "colorir.xaml")

# PARAM
PARAM_REVISOES = "21. NÚMERO DE REVISÕES"

# CONFIG (persistência de estado)
config = script.get_config("PlannixColorir")

# SOFT VARIABLES
STATUSPECA = "18. STATUS DA PEÇA"

statuspj = "Projetada"
statuspg = "Programada"
statuscd = "Corte e Dobra Realizado"
statusar = "Armação Realizada"
statusfo = "Forma Realizada"
statusfa = "Forma com Armação Realizada"
statusco = "Concretagem Realizada"
statuspr = "Preparação Realizada"
statusct = "Corte Realizado"
statuspm = "Pré-montagem Realizada"
statusmm = "Montagem Realizada (Met.)"
statussd = "Solda Realizada"
statusam = "Acabamento Realizado (Met.)"
statusjt = "Jateamento Realizado"
statusgv = "Galvanização Realizada"
statuspt = "Pintura Realizada"
statusac = "Acabamento Realizado"
statusex = "Expedida para a Obra"
statusdv = "Devolvida pela Obra"
statusdc = "Descarregada na Obra"
statusmt = "Montada na Obra"

cordf = Color(192, 192, 192)
corpj = Color(178, 102, 255)
corpg = Color(255, 102, 178)
corcd = Color(255, 185, 185)
corar = Color(255, 145, 145)
corfo = Color(255, 105, 105)
corfa = Color(255,  65,  65)
corco = Color(255,   5,   5)
corpr = Color(255, 205, 205)
corct = Color(255, 180, 180)
corpm = Color(255, 155, 155)
cormm = Color(255, 130, 130)
corsd = Color(255, 105, 105)
coram = Color(255,  80,  80)
corjt = Color(255,  55,  55)
corgv = Color(255,  30,  30)
corpt = Color(255,   5,   5)
corac = Color(255, 178, 102)
corex = Color(255, 128,   0)
cordv = Color(102,   0,   0)
cordc = Color(255, 255, 102)
cormt = Color(102, 255, 102)

# CORES DE REVISÃO
rev0 = Color(192, 192, 192)
rev1 = Color( 0,   191, 255)
rev2 = Color(0, 250,  154)
rev3 = Color(238, 238,   0)
rev4 = Color(249, 178,   8)
rev5 = Color(238,   44,   44)
rev6 = Color(180,   0,   0)

# STATUS → COR
status_color_map = {
    statuspj: corpj,
    statuspg: corpg,
    statuscd: corcd,
    statusar: corar,
    statusfo: corfo,
    statusfa: corfa,
    statusco: corco,
    statuspr: corpr,
    statusct: corct,
    statuspm: corpm,
    statusmm: cormm,
    statussd: corsd,
    statusam: coram,
    statusjt: corjt,
    statusgv: corgv,
    statuspt: corpt,
    statusac: corac,
    statusex: corex,
    statusdv: cordv,
    statusdc: cordc,
    statusmt: cormt,
}

# CATEGORIAS DE INTERESSE
categorias_interesse = [
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Assemblies,
]
categorias_ids = set(int(c) for c in categorias_interesse)

categorias_principais_ids = set([
    int(BuiltInCategory.OST_StructuralColumns),
    int(BuiltInCategory.OST_StructuralFraming),
    int(BuiltInCategory.OST_StructuralFoundation),
    int(BuiltInCategory.OST_Walls),
    int(BuiltInCategory.OST_Floors),
])

# XAML
default_xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Colorir Modelo"
        Height="240"
        Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">

    <StackPanel Margin="15">

        <TextBlock Text="Coloração do Modelo"
                   FontWeight="Bold"
                   FontSize="14"
                   Margin="0,0,0,15"/>

        <RadioButton Name="radio_atualizar"
                     Content="Aplicar cores dos status do Plannix nos elementos"
                     IsChecked="True"
                     Margin="0,0,0,8"/>

        <RadioButton Name="radio_revisoes"
                     Content="Aplicar cores dos números de revisões nos elementos"
                     Margin="0,0,0,8"/>

        <RadioButton Name="radio_remover"
                     Content="Remover cores dos elementos"
                     Margin="0,0,0,15"/>

        <StackPanel Orientation="Horizontal"
                    HorizontalAlignment="Right">

            <Button Name="btn_cancel"
                    Content="Cancelar"
                    Width="90"
                    Margin="0,0,10,0"/>

            <Button Name="btn_ok"
                    Content="Aplicar"
                    Width="90"/>

        </StackPanel>

    </StackPanel>

</Window>
"""

# FUNCTIONS

def get_solid_fill_id():
    patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
    for fp in patterns:
        if fp.GetFillPattern().IsSolidFill:
            return fp.Id
    return ElementId.InvalidElementId


def get_status(element):
    if element is None:
        return ""
    param = element.LookupParameter(STATUSPECA)
    if not param:
        return ""
    value = param.AsValueString() or param.AsString() or ""
    return value.strip()


def get_color_for_status(status_value):
    if not status_value:
        return cordf
    status_lower = status_value.strip().lower()
    for key, color in status_color_map.items():
        if key.lower() == status_lower:
            return color
    return cordf


def get_color_for_revisoes(element):
    param = element.LookupParameter(PARAM_REVISOES)
    if not param:
        return rev0
    try:
        valor = param.AsInteger()
    except:
        return rev0
    if valor <= 0:
        return rev0
    elif valor == 1:
        return rev1
    elif valor == 2:
        return rev2
    elif valor == 3:
        return rev3
    elif valor == 4:
        return rev4
    elif valor == 5:
        return rev5
    else:
        return rev6


def make_override(color, solid_fill_id):
    ogs = OverrideGraphicSettings()
    ogs.SetSurfaceForegroundPatternColor(color)
    if solid_fill_id != ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
    ogs.SetSurfaceTransparency(0)
    return ogs


def get_main_element(assembly):
    for mid in assembly.GetMemberIds():
        membro = doc.GetElement(mid)
        if not membro or not membro.Category:
            continue
        if membro.Category.Id.IntegerValue in categorias_principais_ids:
            return membro
    return None


def apply_colors_generic(get_color_func, transaction_name, msg):
    solid_fill_id = get_solid_fill_id()
    count = 0
    with Transaction(doc, transaction_name) as t:
        t.Start()
        for cat in categorias_interesse:
            if cat == BuiltInCategory.OST_Assemblies:
                continue
            collector = (
                FilteredElementCollector(doc, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            )
            for element in collector:
                color = get_color_func(element)
                ogs   = make_override(color, solid_fill_id)
                view.SetElementOverrides(element.Id, ogs)
                count += 1
        assemblies = (
            FilteredElementCollector(doc, view.Id)
            .OfCategory(BuiltInCategory.OST_Assemblies)
            .WhereElementIsNotElementType()
        )
        for assembly in assemblies:
            if not isinstance(assembly, AssemblyInstance):
                continue
            main_el = get_main_element(assembly)
            color   = get_color_func(main_el)
            ogs     = make_override(color, solid_fill_id)
            view.SetElementOverrides(assembly.Id, ogs)
            count += 1
        t.Commit()
    print("{} {} elementos.".format(msg, count))


def apply_colors():
    apply_colors_generic(
        lambda el: get_color_for_status(get_status(el)),
        "Colorir Modelo - Status",
        "Cores de status aplicadas a"
    )


def apply_revision_colors():
    apply_colors_generic(
        get_color_for_revisoes,
        "Colorir Modelo - Revisões",
        "Cores de revisão aplicadas a"
    )


def remove_colors():
    count = 0
    empty_ogs = OverrideGraphicSettings()
    with Transaction(doc, "Colorir Modelo - Remover") as t:
        t.Start()
        for cat in categorias_interesse:
            collector = (
                FilteredElementCollector(doc, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            )
            for element in collector:
                view.SetElementOverrides(element.Id, empty_ogs)
                count += 1
        t.Commit()
    print("Cores removidas de {} elementos.".format(count))


# MAIN

# Sempre atualizar o XAML
with io.open(xaml_path, "w", encoding="utf-8") as f:
    f.write(default_xaml)

window = forms.WPFWindow(xaml_path)

# Carregar estado salvo
opcao_salva = getattr(config, "opcao_colorir", "atualizar")
if opcao_salva == "remover":
    window.radio_remover.IsChecked = True
elif opcao_salva == "revisoes":
    window.radio_revisoes.IsChecked = True
else:
    window.radio_atualizar.IsChecked = True


def aplicar(sender, args):
    atualizar = window.radio_atualizar.IsChecked
    revisoes  = window.radio_revisoes.IsChecked
    remover   = window.radio_remover.IsChecked

    # Salvar estado
    if remover:
        config.opcao_colorir = "remover"
    elif revisoes:
        config.opcao_colorir = "revisoes"
    else:
        config.opcao_colorir = "atualizar"
    script.save_config()

    window.Close()
    try:
        if atualizar:
            apply_colors()
        elif revisoes:
            apply_revision_colors()
        elif remover:
            remove_colors()
    except Exception as e:
        error_msg = str(e).lower()
        if "central" in error_msg or "network" in error_msg or "workset" in error_msg:
            print(
                "O modelo central está inacessível. "
                "Verifique a conexão com o servidor e tente novamente."
            )
        else:
            print("Erro inesperado: {}".format(str(e)))


def cancelar(sender, args):
    window.Close()


window.btn_ok.Click     += aplicar
window.btn_cancel.Click += cancelar

window.show_dialog()