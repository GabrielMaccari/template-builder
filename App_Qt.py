# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
import platform

import PyQt6.QtWidgets as QtWidgets

import ViewController_Qt

OS = platform.platform()


def main():
    app = QtWidgets.QApplication(sys.argv)

    if OS.startswith("Win"):
        app.setStyle("fusion")
    else:
        app.setStyle("Breeze")

    janela = ViewController_Qt.JanelaPrincipalApp()
    janela.show()
    app.exec()


if __name__ == '__main__':
    main()
