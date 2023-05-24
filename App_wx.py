# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""
import wx
import ctypes

import ViewController_wx


def main():
	try:
		ctypes.windll.shcore.SetProcessDpiAwareness(True)
	except Exception as e:
		print(f"{e.__class__}: {e}")

	app = wx.App()

	janela = ViewController_wx.JanelaPrincipalApp()
	janela.Centre()
	janela.Show()

	app.MainLoop()


if __name__ == '__main__':
	main()
