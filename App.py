# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from PyQt6.QtWidgets import QApplication

from View import JanelaPrincipalApp


if __name__ == '__main__':
    app = QApplication(sys.argv)

    with open("recursos_app/estilos/estilo.qss", 'r') as f:
        estilo = f.read()
    app.setStyleSheet(estilo)

    window = JanelaPrincipalApp()
    window.show()
    app.exec()
