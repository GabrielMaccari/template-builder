# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""
import wx
import ctypes

from ViewController_wx import JanelaPrincipalApp

if __name__ == '__main__':
	try:
		ctypes.windll.shcore.SetProcessDpiAwareness(True)
	except:
		pass

	app = wx.App()

	janela = JanelaPrincipalApp()
	janela.Centre()
	janela.Show()

	app.MainLoop()
