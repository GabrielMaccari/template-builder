# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from PyQt6.QtWidgets import QApplication

from View_Qt import JanelaPrincipalApp


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("fusion")

    window = JanelaPrincipalApp()
    window.show()
    app.exec()
