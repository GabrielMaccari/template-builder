# -*- coding: utf-8 -*-
""" @author: Gabriel Maccari """

import sys
from platform import platform
from PyQt6.QtWidgets import QApplication
from docx.opc.exceptions import PackageNotFoundError
from icecream import ic

from Model import Modelo
from View import Interface
from Controller import Controlador, mostrar_popup

ic.configureOutput(prefix='LOG| ', includeContext=True)

OS = platform()

TEMPLATE_ESTILOS = "config/template_estilos.docx"
JSON_COLUNAS = "config/colunas_aba_geral.json"


def erro_dependencia(arquivo, excecao):
    ic(excecao)
    mostrar_popup(
        f"Dependência não encontrada: {arquivo}. Restaure o arquivo a partir do repositório e tente novamente.",
        tipo_msg="erro",
    )
    sys.exit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion" if OS.startswith("Win") else "Breeze")

    try:
        # Checa se o arquivo de definição das colunas existe
        with open(JSON_COLUNAS, 'r'):
            pass

        # Inicializa os componentes
        model = Modelo(TEMPLATE_ESTILOS, JSON_COLUNAS)
        view = Interface(model.colunas)
        Controlador(model, view)

    except PackageNotFoundError as erro:
        erro_dependencia(TEMPLATE_ESTILOS, erro)

    except FileNotFoundError as erro:
        erro_dependencia(JSON_COLUNAS, erro)

    app.exec()
