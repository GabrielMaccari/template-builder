# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys
from PyQt6.QtWidgets import *
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
from docx.opc.exceptions import PackageNotFoundError

from Controller import ControladorPrincipal, COLUNAS_TABELA_CADERNETA


class JanelaPrincipalApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Carrega o template de estilos da caderneta e instancia o controlador
        template = f"recursos_app/modelos/template_estilos.docx"
        try:
            self.controlador = ControladorPrincipal(template)
        except PackageNotFoundError:
            mostrar_popup(
                f"Dependência não encontrada: {template}. Restaure o arquivo a "
                f"partir do repositório e tente novamente.",
                tipo_msg="erro", parent=self
            )
            sys.exit()

        # Constrói a interface -----
        self.setWindowTitle('Template Builder')
        self.setWindowIcon(QIcon('recursos_app/icones/book.png'))

        # Rótulo do nome do arquivo
        self.rotulo_arquivo = QLabel("Selecione um arquivo .xlsx com os dados dos pontos mapeados.")

        # Botão para selecionar a tabela
        botao_abrir_arquivo = QPushButton("Selecionar")
        botao_abrir_arquivo.setMaximumWidth(75)
        botao_abrir_arquivo.clicked.connect(self.botao_abrir_arquivo_clicado)

        # Layout que organiza a porção superior da interface (acima da linha)
        layout_superior = QHBoxLayout()

        layout_superior.addWidget(self.rotulo_arquivo)
        layout_superior.addWidget(botao_abrir_arquivo)

        # Linha que separa a seleção de arquivo do restante da interface
        separador = QFrame(None)
        separador.setLineWidth(1)
        separador.setFrameShape(QFrame.Shape.HLine)
        separador.setFrameShadow(QFrame.Shadow.Sunken)

        # Layout em grade da seção de colunas da tabela
        layout_central = QGridLayout(None)
        layout_central.setHorizontalSpacing(0)
        layout_central.setVerticalSpacing(5)
        layout_central.setColumnMinimumWidth(2, 50)

        # Cria os rótulos das colunas e botões de status em listas e adiciona-os
        # ao layout em grade da seção central
        self.rotulos_colunas, self.botoes_status = [], []
        lin, col = 0, 0
        for i, coluna in enumerate(COLUNAS_TABELA_CADERNETA):
            self.rotulos_colunas.append(QLabel(coluna))
            self.botoes_status.append(BotaoStatus(coluna, self))

            layout_central.addWidget(self.rotulos_colunas[i], lin, col)
            layout_central.addWidget(self.botoes_status[i], lin, col + 1)

            col = col if lin != 8 else col + 3
            lin = lin + 1 if lin != 8 else 0

        # Rótulos para exibir o número de linhas (pontos meapeados) na tabela
        rotulo_num_pontos = QLabel("Número de pontos na tabela: ")
        self.num_pontos = QLabel("-")
        self.num_pontos.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout_central.addWidget(rotulo_num_pontos, 9, 0)
        layout_central.addWidget(self.num_pontos, 9, 1)

        # Linha que separa a porção central da porção inferior da interface
        separador2 = QFrame(None)
        separador2.setLineWidth(1)
        separador2.setFrameShape(QFrame.Shape.HLine)
        separador2.setFrameShadow(QFrame.Shadow.Sunken)

        # Checkbox para marcar se gera ou não a folha de rosto
        self.checkbox_folha_rosto = QCheckBox("Incluir folha de rosto no início da caderneta")
        self.checkbox_folha_rosto.setToolTip("Gera uma página com informações do projeto no início da caderneta")
        self.checkbox_folha_rosto.setChecked(True)

        # Botão para gerar e exportar o template da caderneta
        self.botao_gerar_modelo = QPushButton("Gerar caderneta")
        self.botao_gerar_modelo.setMinimumHeight(35)
        self.botao_gerar_modelo.clicked.connect(self.botao_gerar_caderneta_clicado)
        self.botao_gerar_modelo.setEnabled(False)

        # Layout mestre (aninha os widgets e demais layouts)
        layout_principal = QVBoxLayout()
        layout_principal.setSpacing(5)

        layout_principal.addLayout(layout_superior)
        layout_principal.addWidget(separador)
        layout_principal.addLayout(layout_central)
        layout_principal.addWidget(separador2)
        layout_principal.addWidget(self.checkbox_folha_rosto)
        layout_principal.addWidget(self.botao_gerar_modelo)

        # Painel contendo o layout
        container = QWidget(self)
        container.setLayout(layout_principal)
        self.setCentralWidget(container)

    def botao_abrir_arquivo_clicado(self):
        """
        Chama a função para abrir a tabela no controlador. Caso um arquivo seja aberto com sucesso, atualiza o
        rotulo_arquivo com o nome do arquivo selecionado e chama o método checar_colunas para atualizar os ícones de
        status.
        :returns: Nada.
        """
        try:
            arquivo_aberto, num_pontos = False, "-"
            caminho = mostrar_dialogo_arquivo(
                "Selecione uma tabela contendo os dados de entrada.",
                "Pasta de trabalho do Excel (*.xlsx);;Pasta de trabalho habilitada para macro do Excel (*.xlsm);;"
            )
            if caminho != "":
                arquivo_aberto, num_pontos = self.controlador.abrir_tabela(caminho)
            if arquivo_aberto:
                partes_caminho = caminho.split("/")
                self.rotulo_arquivo.setText(partes_caminho[-1])
                self.num_pontos.setText(str(num_pontos) if num_pontos > 0 else "-")
                self.checar_colunas()

                if "nan" in self.controlador.df.columns:
                    mostrar_popup("Atenção! Existem colunas com nomes inválidos na tabela que podem causar erros ou "
                                  "anomalias no funcionamento da ferramenta. Verifique se as fórmulas presentes nas "
                                  "células de cabeçalho das colunas de estruturas (colunas S a AG) não foram "
                                  "comprometidas. Isso geralmente ocorre ao recortar e colar células na aba de Listas "
                                  "ao preencher as estruturas.")

        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro")
            print(exception.__class__, exception)

    def checar_colunas(self):
        """
        Chama a função do controlador para checar se cada coluna está no formato especificado. Atualiza os botões de
        status conforme o resultado. Se todas as colunas estiverem OK, habilita o botão para gerar o template.
        :returns: Nada.
        """
        status_colunas = self.controlador.checar_colunas()
        for widget, status in zip(self.botoes_status, status_colunas):
            widget.definir_status(status)
        if all(stts == "ok" for stts in status_colunas):
            self.botao_gerar_modelo.setEnabled(True)
        else:
            self.botao_gerar_modelo.setEnabled(False)

    def icone_status_clicado(self, coluna: str, status: str):
        """
        Chama a função do controlador para identificar os problemas na coluna e mostra os resultados em uma popup.
        :param coluna: A coluna a ser verificada.
        :param status: O status da coluna ("ok", "faltando", "problemas", "nulos" ou "dominio").
        :returns: Nada.
        """

        try:
            localizar_problemas = {
                "missing_column": lambda coluna_faltando: [],
                "wrong_dtype": self.controlador.localizar_problemas_formato,
                "nan_not_allowed": self.controlador.localizar_celulas_vazias,
                "outside_domain": self.controlador.localizar_problemas_dominio
            }

            if status not in localizar_problemas.keys():
                raise Exception(f"O status informado não foi reconhecido: {status}")

            indices_problemas = localizar_problemas[status](coluna)

            msg = self.controlador.montar_msg_problemas(status, coluna, indices_problemas)
            mostrar_popup(msg, "notificacao", self)

        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro")

    def botao_gerar_caderneta_clicado(self):
        """
        Chama as funções do controlador para gerar a caderneta e exportá-la.
        :returns: Nada.
        """
        try:
            mostrar_cursor_espera()
            montar_folha_de_rosto = self.checkbox_folha_rosto.isChecked()
            self.controlador.gerar_caderneta(montar_folha_de_rosto)
            mostrar_cursor_espera(False)
            caminho = mostrar_dialogo_arquivo("Salvar documento da caderneta", "*.docx", modo="salvar")
            if caminho != "":
                self.controlador.salvar_caderneta(caminho)
                mostrar_popup("Caderneta criada com sucesso!")
        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro")


class BotaoStatus(QPushButton):
    def __init__(self, coluna: str, parent: QMainWindow, status: str = "none"):
        super().__init__()
        self.coluna = coluna
        self.parent = parent
        self.status = status
        self.setMaximumSize(18, 18)
        self.setFlat(True)
        self.setStyleSheet(
            """
            background-color: none;
            border-color: none;
            border-width: 0;
            """
        )
        self.definir_status(status)

    def definir_status(self, status: str):
        """
        Define o ícone e a tooltip do botão. Conecta à função icone_status_clicado caso o status não seja "none" ou "ok".
        :param status: O status da coluna ("none", "ok", "faltando", "problemas", "nulos" ou "dominio")
        :returns: Nada.
        """
        dic_botoes = {
            "none": {
                "icone": QIcon("recursos_app/icones/circle.png"),
                "tooltip": "Carregue um arquivo"
            },
            "ok": {
                "icone": QIcon("recursos_app/icones/ok.png"),
                "tooltip": "OK"
            },
            "missing_column": {
                "icone": QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "Coluna não encontrada na tabela"
            },
            "wrong_dtype": {
                "icone": QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna contém dados com\nformato errado"
            },
            "nan_not_allowed": {
                "icone": QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna não permite nulos,\nmas existem células vazias"
            },
            "outside_domain": {
                "icone": QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "Algumas células contêm valores\nfora da lista de valores permitidos"
            }
        }

        icone = QIcon(dic_botoes[status]["icone"])
        tooltip = dic_botoes[status]["tooltip"]

        self.status = status
        self.setIcon(icone)
        self.setToolTip(tooltip)

        try:
            self.clicked.disconnect()
        except TypeError:
            pass

        if status not in ["none", "ok"]:
            self.clicked.connect(lambda: self.parent.icone_status_clicado(self.coluna, self.status))

        self.setEnabled(status != "none")


def mostrar_dialogo_arquivo(titulo: str, filtro: str, modo="abrir", parent: QMainWindow = None):
    """
    Abre um diálogo de seleção/salvamento de arquivo.
    :param titulo: O título da janela.
    :param filtro: Filtros de tipo de arquivo (Ex: "Planilha do Excel (*.xlsx);;Planilha com macro do Excel (*.xlsm)")
    :param modo: "abrir" ou "salvar". Define se o diálogo será de abertura ou salvamento de arquivo.
    :param parent: A janela pai (Default = None).
    :returns: Nada.
    """
    dialog = QFileDialog(parent)
    if modo == "abrir":
        caminho, tipo = dialog.getOpenFileName(
            caption=titulo, filter=filtro, parent=parent
        )
    else:
        caminho, tipo = dialog.getSaveFileName(
            caption=titulo, filter=filtro, parent=parent
        )
    return caminho


def mostrar_cursor_espera(ativar: bool = True):
    """
    Troca o cursor do mouse por um cursor de espera.
    :param ativar: Default True. False para restaurar o cursor normal.
    :returns: Nada.
    """
    if ativar:
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
    else:
        QApplication.restoreOverrideCursor()


def mostrar_popup(mensagem: str, tipo_msg: str = "notificacao", parent: QMainWindow = None):
    """
    Mostra uma popup com uma mensagem ao usuário.
    :param mensagem: A mensagem a ser exibida na popup.
    :param tipo_msg: "notificacao" ou "erro" (define o ícone da popup). O valor padrão é "notificacao".
    :param parent: A janela pai (Default = None).
    :returns: Nada.
    """
    tipos_popup = {
        "notificacao": {"titulo": "Notificação", "icone": "recursos_app/icones/info.png"},
        "erro":        {"titulo": "Erro",        "icone": "recursos_app/icones/error.png"}
    }
    title = tipos_popup[tipo_msg]["titulo"]
    icon = QIcon(tipos_popup[tipo_msg]["icone"])

    popup = QMessageBox(parent)
    popup.setText(mensagem)
    popup.setWindowTitle(title)
    popup.setWindowIcon(icon)
    popup.exec()
