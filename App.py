# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from platform import platform
from PyQt6.QtWidgets import QApplication
from docx.opc.exceptions import PackageNotFoundError

from Model import Modelo
from View import Interface
from Controller import Controlador, mostrar_popup

OS = platform()
TEMPLATE = "recursos_app/modelos/template_estilos.docx"

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion" if OS.startswith("Win") else "Breeze")

    # Carrega o template de estilos da caderneta e instancia o controlador
    try:
        model = Modelo(TEMPLATE)
        view = Interface()
        controller = Controlador(model, view)
    except PackageNotFoundError:
        mostrar_popup(
            f"Dependência não encontrada: {TEMPLATE}. Restaure o arquivo a "
            f"partir do repositório e tente novamente.",
            tipo_msg="erro",
        )
        sys.exit()

    app.exec()
