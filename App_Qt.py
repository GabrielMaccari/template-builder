# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from platform import platform
from PyQt6.QtWidgets import QApplication

from ViewController_Qt import JanelaPrincipalApp

OS = platform()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    if OS.startswith("Win"):
        app.setStyle("fusion")
    else:
        app.setStyle("Breeze")

    janela = JanelaPrincipalApp()
    janela.show()
    app.exec()
