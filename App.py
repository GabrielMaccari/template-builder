# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from platform import platform
from PyQt6.QtWidgets import QApplication
from docx.opc.exceptions import PackageNotFoundError
from icecream import ic

from Model import Modelo
from View import Interface
from Controller import Controlador, mostrar_popup

ic.configureOutput(prefix='LOG| ', includeContext=True)
"""
# Configuração para geração de logs de execução (Obs: NÃO FUNCIONA COM PYINSTALLER!!!)
def log_to_file(text, mode='a'):
    with open("log.log", mode, encoding="utf-8") as f:
        f.write(f"{text}\n")
ic.configureOutput(prefix='LOG| ', includeContext=True, outputFunction=log_to_file)
log_to_file("------------------ LOG DA ÚLTIMA EXECUÇÃO ------------------", 'w')
# ------------------------------------------------------------------------------------
"""

OS = platform()
TEMPLATE = "config/modelos/template_estilos.docx"

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
