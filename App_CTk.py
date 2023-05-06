# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import customtkinter as ctk

from View_CTk import AppMainWindow


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("dark-blue")
    view = AppMainWindow()
    view.mainloop()
