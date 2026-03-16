# -*- coding: utf-8 -*-

# HEADER
__title__ = "Configurações"
__author__ = "Daniel Viegas"
__doc__ = """Configure as opções desejadas para a exportação de dados para a integração entre o software Revit e o software Plannix."""

# IMPORTS
import os
import io
from pyrevit import forms, script

output = script.get_output()

try:

    # HARD VARIABLES
    config = script.get_config("PlannixProject")
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, "config.xaml")

    # DEFAULT XAML
    default_xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Configurações"
        Height="240"
        Width="460"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">

    <StackPanel Margin="15">

        <TextBlock Text="Configurações de Exportação"
                   FontWeight="Bold"
                   FontSize="14"
                   Margin="0,0,0,15"/>

        <CheckBox Name="check_print_pdf" Margin="0,0,0,10">
            <TextBlock Text="Imprimir PDFs de detalhamento ao exportar XML?"
                       TextWrapping="Wrap"/>
        </CheckBox>

        <CheckBox Name="check_overwrite_pdf" Margin="0,0,0,20">
            <TextBlock Text="Sobrescrever arquivos com mesmo nome ao imprimir PDFs?"
                       TextWrapping="Wrap"/>
        </CheckBox>

        <StackPanel Orientation="Horizontal"
                    HorizontalAlignment="Right">

            <Button Name="btn_cancel"
                    Content="Cancelar"
                    Width="90"
                    Margin="0,0,10,0"/>

            <Button Name="btn_ok"
                    Content="Salvar"
                    Width="90"/>

        </StackPanel>

    </StackPanel>

</Window>
"""

    # Criar XAML se não existir
    if not os.path.exists(xaml_path):
        with io.open(xaml_path, "w", encoding="utf-8") as f:
            f.write(default_xaml)

    # Abrir janela
    window = forms.WPFWindow(xaml_path)

    # Carregar configurações
    window.check_print_pdf.IsChecked = getattr(config, "print_pdfs", True)
    window.check_overwrite_pdf.IsChecked = getattr(config, "overwrite_pdfs", False)

    # Evento dos botões
    def salvar(sender, args):

        config.print_pdfs = window.check_print_pdf.IsChecked
        config.overwrite_pdfs = window.check_overwrite_pdf.IsChecked

        script.save_config()

        window.Close()


    def cancelar(sender, args):
        window.Close()


    window.btn_ok.Click += salvar
    window.btn_cancel.Click += cancelar

    # Mostrar janela
    window.show_dialog()


finally:
    output.close()