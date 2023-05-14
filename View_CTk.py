# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""
import sys
import customtkinter as ctk
from tktooltip import ToolTip
from docx.opc.exceptions import PackageNotFoundError
from PIL import Image, ImageEnhance
from tkinter import filedialog

from Controller import ControladorPrincipal, COLUNAS_TABELA_CADERNETA

OS = sys.platform
THEME = ctk.get_appearance_mode()
COLORS = {
    "background": ("#FBFBFB","#242424"),
    "background_hover": ("#F6F6F6","#313131"),
    "border": ("#d9d9d9","#1D1D1D"),
    "scrollbar": ("#AFAFAF","#9A9A9A"),
    "tooltip": ("#ECECEC","#2D2D2D"),
    "text": ("#1B1212","#BCBEC4"),
    "button": ("#0e6aba","#313131"),
    "button_hover": ("#0a4c86","#363636")
}
FONT = ("Cantarell", 11) if OS == "linux" else ("Segoe UI", 12)


class AppMainWindow(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Carrega o template de estilos da caderneta e instancia o controlador
        template = f"recursos_app/modelos/template_estilos.docx"
        try:
            self.controller = ControladorPrincipal(template)
        except PackageNotFoundError:
            MessagePopup(f"Dependência não encontrada: {template}. Restaure o arquivo "
                         f"a partir do repositório e tente novamente.", self)
            sys.exit()

        self.title("Template Builder")
        if OS == "win32": self.iconbitmap("recursos_app/icones/book.ico")
        self.configure(fg_color=("#FBFBFB","#242424"))

        main_layout = ctk.CTkFrame(self, fg_color=COLORS["background"])
        main_layout.pack(padx=5, pady=5, fill=ctk.BOTH)

        upper_layout = ctk.CTkFrame(main_layout, fg_color=COLORS["background"])
        upper_layout.columnconfigure(0, weight=4)
        upper_layout.columnconfigure(1, weight=1)

        self.file_label = Label(upper_layout, text="Selecione um arquivo .xlsx com os dados dos pontos mapeados.")
        self.file_label.grid(row=0, column=0, columnspan=3, padx=(0,5), sticky=ctk.NSEW)

        self.open_file_button = ctk.CTkButton(
            upper_layout, text="Selecionar", font=FONT, width=70, height=25, border_width=0,
            corner_radius=3, text_color="white", fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"], command=self.open_file_button_clicked
        )
        self.open_file_button.grid(row=0, column=3, sticky=ctk.E)

        upper_layout.pack(fill=ctk.X)

        separator = ctk.CTkFrame(main_layout, height=1, border_color=COLORS["border"], border_width=2)
        separator.pack(fill=ctk.X, pady=(5,5))

        central_layout = ctk.CTkFrame(main_layout, fg_color=COLORS["background"])
        central_layout.columnconfigure(0, weight=3)
        central_layout.columnconfigure(1, weight=1)
        central_layout.columnconfigure(3, weight=3)
        central_layout.columnconfigure(4, weight=1)

        # Cria os rótulos das columns e botões de status em listas e adiciona-os ao layout em grade da seção central
        self.column_labels, self.status_buttons = [], []
        r, c = 0, 0
        for i, column in enumerate(COLUNAS_TABELA_CADERNETA):
            self.column_labels.append(Label(central_layout, column))
            self.status_buttons.append(BotaoStatus(column, central_layout, self))
            self.column_labels[i].grid(row=r, column=c, sticky=ctk.W)
            self.status_buttons[i].grid(row=r, column=c+1, padx=(5,0))
            c = c if r != 8 else c + 3
            r = r + 1 if r != 8 else 0

        spacer = ctk.CTkFrame(central_layout, width=50, fg_color=COLORS["background"])
        spacer.grid(row=0, column=2, rowspan=9, sticky="WE")

        point_number_label = Label(central_layout, "Número de pontos na tabela:")
        point_number_label.grid(row=9, column=0, sticky=ctk.W)
        self.number_of_points = Label(central_layout, "-")
        self.number_of_points.grid(row=9, column=1, padx=(5,0))

        central_layout.pack(fill=ctk.X)

        separator2 = ctk.CTkFrame(main_layout, height=1, border_color=COLORS["border"], border_width=2)
        separator2.pack(fill=ctk.X, pady=(5, 5))

        self.title_page_checkbox = ctk.CTkCheckBox(
            main_layout, text="Incluir folha de rosto no início da caderneta", onvalue=True, offvalue=False, checkbox_width=12,
            checkbox_height=12, corner_radius=2, border_width=2, font=FONT, text_color=COLORS["text"],
            fg_color=COLORS["button"], hover_color=COLORS["button_hover"]
        )
        self.title_page_checkbox.pack(fill=ctk.X, anchor=ctk.W)
        self.title_page_checkbox.select()

        self.fill_template_button = ctk.CTkButton(
            main_layout, text="Gerar caderneta", height=35, corner_radius=3, border_width=0,
            fg_color=COLORS["button"], hover_color=COLORS["button_hover"], text_color="white", font=FONT
        )
        self.fill_template_button.pack(fill=ctk.X, pady=(5,0))
        self.fill_template_button.configure(state=ctk.DISABLED, command=self.fill_template_button_clicked)

        self.resizable(False, False)

    def open_file_button_clicked(self):
        try:
            file_open, num_points = False, "-"
            path = filedialog.askopenfilename(
                title="Selecione uma tabela de pontos",
                filetypes=(("Pasta de trabalho do Excel", "*.xlsx"),
                           ("Pasta de trabalho habilitada para macro do Excel", "*.xlsm")))
            if path != "":
                file_open, num_points = self.controller.abrir_tabela(path)
            if file_open:
                split_path = path.split("/")
                self.file_label.configure(text=split_path[-1])
                self.number_of_points.configure(text=str(num_points) if num_points > 0 else "-")
                self.check_column_status()

                if "nan" in self.controller.df.columns:
                    MessagePopup("Atenção! Existem colunas com nomes inválidos na tabela que podem causar erros ou "
                                 "anomalias no funcionamento da ferramenta. Verifique se as fórmulas presentes nas "
                                 "células de cabeçalho das colunas de estruturas (colunas S a AG) não foram "
                                 "comprometidas. Isso geralmente ocorre ao recortar e colar células na aba de Listas "
                                 "ao preencher as estruturas.", self)

        except Exception as exception:
            MessagePopup(f"ERRO: {exception}", self)

    def check_column_status(self):
        column_status = self.controller.checar_colunas()
        for stts_btn, status in zip(self.status_buttons, column_status):
            stts_btn.set_status(status)
        all_ok = True if all(stts == "ok" for stts in column_status) else False
        self.fill_template_button.configure(state=ctk.NORMAL if all_ok else ctk.DISABLED)

    def status_button_clicked(self, column: str, status: str):
        try:
            search_problems = {
                "missing_column": lambda missing_column: [],
                "wrong_dtype": self.controller.localizar_problemas_formato,
                "nan_not_allowed": self.controller.localizar_celulas_vazias,
                "outside_domain": self.controller.localizar_problemas_dominio
            }

            if status not in search_problems.keys():
                raise Exception(f"O status informado não foi reconhecido: {status}")

            problem_indexes = search_problems[status](column)

            msg = self.controller.montar_msg_problemas(status, column, problem_indexes)
            MessagePopup(msg, self)
        except Exception as exception:
            MessagePopup(f"ERRO: {exception}", self)

    def fill_template_button_clicked(self):
        try:
            self.show_wait_cursor()
            include_title_page = self.title_page_checkbox.get()
            self.controller.gerar_caderneta(include_title_page)
            self.show_wait_cursor(False)
            path = filedialog.asksaveasfilename(title="Salvar caderneta", defaultextension="docx",
                                                filetypes=(("Documento do Word", "*.docx"), ("Documento do Word", "*.docx")))
            if path != "":
                self.controller.salvar_caderneta(path)
                MessagePopup("Caderneta criada com sucesso!", self)
        except Exception as exception:
            MessagePopup(f"ERRO: {exception}", self)

    def show_wait_cursor(self, activate: bool = True):
        self.config(cursor="wait" if activate else "")
        self.update()


class Label(ctk.CTkLabel):
    def __init__(self, master, text, height=18):
        super().__init__(master, text=text, font=FONT, height=height, anchor=ctk.W)


class BotaoStatus(ctk.CTkButton):
    def __init__(self, column: str, master: ctk.CTkFrame, parent: ctk.CTk, status: str = "none"):
        super().__init__(master=master, width=20, height=20, border_width=0, corner_radius=4,
                         text="", fg_color=COLORS["background"], hover_color=COLORS["background_hover"],
                         image=ctk.CTkImage(Image.open("recursos_app/icones/circle.png"), size=(16, 16)))
        self.column = column
        self.parent = parent
        self.status = status

        self.icon = None
        self.tooltip = None

        self.set_status(status)

    def set_status(self, status):
        """
                Define o ícone e a tooltip do botão. Conecta à função icone_status_clicado caso o status não seja "none" ou "ok".
                :param status: O status da column ("none", "ok", "faltando", "problemas", "nulos" ou "dominio")
                :returns: Nada.
                """
        icon_tooltip_dict = {
            "none": {
                "icone": Image.open("recursos_app/icones/circle.png"),
                "tooltip": "Carregue um arquivo"
            },
            "ok": {
                "icone": Image.open("recursos_app/icones/ok.png"),
                "tooltip": "OK"
            },
            "missing_column": {
                "icone": Image.open("recursos_app/icones/not_ok.png"),
                "tooltip": "coluna não encontrada na tabela"
            },
            "wrong_dtype": {
                "icone": Image.open("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna contém dados com\nformato errado"
            },
            "nan_not_allowed": {
                "icone": Image.open("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna não permite nulos,\nmas existem células vazias"
            },
            "outside_domain": {
                "icone": Image.open("recursos_app/icones/not_ok.png"),
                "tooltip": "Algumas células contêm valores\nfora da lista de valores permitidos"
            }
        }

        self.configure(command=None)

        icon = icon_tooltip_dict[status]["icone"]
        self.icon = icon if status != "none" else ImageEnhance.Color(icon).enhance(0)
        self.configure(image=ctk.CTkImage(self.icon, size=(16, 16)))
        self.build_tooltip(icon_tooltip_dict[status]["tooltip"])

        self.status = status

        if status not in ["none", "ok"]:
            self.configure(command=lambda: self.parent.status_button_clicked(self.column, self.status))

    def build_tooltip(self, text):
        if isinstance(self.tooltip, ToolTip):
            try:
                self.tooltip.destroy() # TODO isso gera uns bugs estranhos, mas ainda assim funciona
            except:
                pass
        theme = ctk.get_appearance_mode()
        bg_color = (COLORS["tooltip"][0] if theme == "Light" else COLORS["tooltip"][1])
        text_color = COLORS["text"][0] if theme == "Light" else COLORS["text"][1]
        self.tooltip = ToolTip(self, msg=text, bg=bg_color, fg=text_color, delay=0.5, padx=3, pady=3)


class MessagePopup(ctk.CTkToplevel):
    def __init__(self, msg: str, master: ctk.CTk = None):
        super().__init__(master=master)

        self.title("Notificação")
        self.configure(fg_color=("#FBFBFB", "#242424"))
        self.minsize(250, 70)

        self.message = Label(self, msg)
        self.message.configure(wraplength=400, anchor=ctk.N)
        self.message.pack(fill=ctk.X, padx=10, pady=10)

        self.ok_button = ctk.CTkButton(
            self, text="OK", width=40, height=20, text_color="white", fg_color=COLORS["button"],
            corner_radius=3, hover_color=COLORS["button_hover"], command=self.destroy
        )
        self.ok_button.pack(anchor=ctk.S, padx=10, pady=(0,10))

        x, y = master.winfo_rootx(), master.winfo_rooty()
        geometry = "+%d+%d" % (x + 100, y + 150)
        self.geometry(geometry)

        self.grab_set()
