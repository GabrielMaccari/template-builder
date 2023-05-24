# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

import sys

import PyQt6.QtWidgets as QtWidgets
import PyQt6.QtGui as QtGui
import PyQt6.QtCore as QtCore
from docx.opc import exceptions as docx_exceptions

from Controller import ControladorPrincipal, COLUNAS_TABELA_CADERNETA

TEMPLATE = f"recursos_app/modelos/template_estilos.docx"


class JanelaPrincipalApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(JanelaPrincipalApp, self).__init__(None)

        # Carrega o template de estilos da caderneta e instancia o controlador
        try:
            self.controlador = ControladorPrincipal(TEMPLATE)
        except docx_exceptions.PackageNotFoundError:
            mostrar_popup(
                f"Dependência não encontrada: {TEMPLATE}. Restaure o arquivo a "
                f"partir do repositório e tente novamente.",
                tipo_msg="erro", parent=self
            )
            sys.exit()

        # Constrói a interface -----
        self.setWindowTitle('Template Builder')
        self.setWindowIcon(QtGui.QIcon('recursos_app/icones/book.png'))

        # Layout que organiza a porção superior da interface (acima da linha)
        layout_arquivo = QtWidgets.QHBoxLayout()

        # Rótulo do nome do arquivo
        self.rotulo_arquivo = QtWidgets.QLabel("Selecione uma tabela com os dados dos pontos mapeados.")

        # Botão para selecionar a tabela
        self.botao_abrir_arquivo = QtWidgets.QPushButton("Selecionar")
        self.botao_abrir_arquivo.setMaximumSize(75,30)
        self.botao_abrir_arquivo.clicked.connect(self.botao_abrir_arquivo_clicado)

        layout_arquivo.addWidget(self.rotulo_arquivo)
        layout_arquivo.addWidget(self.botao_abrir_arquivo)

        # Linha que separa a seleção de arquivo do restante da interface
        separador = QtWidgets.QFrame(None)
        separador.setLineWidth(1)
        separador.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separador.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)

        # Layout em grade da seção de colunas da tabela
        layout_colunas = QtWidgets.QGridLayout(None)

        # Cria os rótulos das colunas e botões de status em listas e adiciona-os ao layout em grade da seção central
        self.rotulos_colunas, self.botoes_status = [], []
        for coluna in COLUNAS_TABELA_CADERNETA:
            self.rotulos_colunas.append(QtWidgets.QLabel(coluna))
            self.botoes_status.append(BotaoStatus(coluna, self))

        tamanho_coluna_grid = int(len(self.rotulos_colunas) / 2)
        if len(self.rotulos_colunas) % 2 != 0:
            tamanho_coluna_grid += 1
        i = 0
        while i < tamanho_coluna_grid:
            try:
                layout_colunas.addWidget(self.rotulos_colunas[i], i, 0)
                layout_colunas.addWidget(self.botoes_status[i], i, 1)
                layout_colunas.addWidget(self.rotulos_colunas[i+tamanho_coluna_grid], i, 3)
                layout_colunas.addWidget(self.botoes_status[i+tamanho_coluna_grid], i, 4)
            except IndexError:
                break
            i += 1

        layout_colunas.setHorizontalSpacing(0)
        layout_colunas.setVerticalSpacing(5)
        layout_colunas.setColumnMinimumWidth(2, 20)

        # Layout para organizar os rótulos de contagem de pontos
        layout_num_pontos = QtWidgets.QHBoxLayout()

        # Rótulos para exibir o número de linhas (pontos meapeados) na tabela
        self.rotulo_num_pontos = QtWidgets.QLabel("Número de pontos na tabela:   ")
        self.num_pontos = QtWidgets.QLabel("-")
        self.num_pontos.setMaximumWidth(30)
        self.num_pontos.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)

        layout_num_pontos.addWidget(self.rotulo_num_pontos)
        layout_num_pontos.addWidget(self.num_pontos)
        layout_num_pontos.addStretch()

        # Linha que separa a porção central da porção inferior da interface
        separador2 = QtWidgets.QFrame(None)
        separador2.setLineWidth(1)
        separador2.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separador2.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)

        # Checkbox para marcar se gera ou não a folha de rosto
        self.checkbox_folha_rosto = QtWidgets.QCheckBox("Incluir folha de rosto no início da caderneta")
        self.checkbox_folha_rosto.setToolTip("Gera uma página com informações do projeto no início da caderneta")
        self.checkbox_folha_rosto.setChecked(True)

        # Botão para gerar e exportar o template da caderneta
        self.botao_gerar_modelo = QtWidgets.QPushButton("Gerar caderneta")
        self.botao_gerar_modelo.setMinimumHeight(35)
        self.botao_gerar_modelo.clicked.connect(self.botao_gerar_caderneta_clicado)
        self.botao_gerar_modelo.setEnabled(False)

        # Layout mestre (aninha os widgets e demais layouts)
        layout_principal = QtWidgets.QVBoxLayout()
        layout_principal.setSpacing(5)

        layout_principal.addLayout(layout_arquivo)
        layout_principal.addWidget(separador)
        layout_principal.addLayout(layout_colunas)
        layout_principal.addLayout(layout_num_pontos)
        layout_principal.addWidget(separador2)
        layout_principal.addWidget(self.checkbox_folha_rosto)
        layout_principal.addWidget(self.botao_gerar_modelo)

        # Painel contendo o layout
        container = QtWidgets.QWidget(self)
        container.setLayout(layout_principal)
        self.setCentralWidget(container)

    def botao_abrir_arquivo_clicado(self):
        """Chama a função para abrir a tabela no controlador. Caso um arquivo seja aberto com sucesso, atualiza o
        rotulo_arquivo com o nome do arquivo selecionado e chama o método checar_colunas para atualizar os ícones de
        status.
        :return: Nada.
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
                                  "ao preencher as estruturas.", parent=self)

        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro", parent=self)

    def checar_colunas(self):
        """Chama a função do controlador para checar se cada coluna está no formato especificado. Atualiza os botões de
        status conforme o resultado. Se todas as colunas estiverem OK, habilita o botão para gerar o template.
        :return: Nada.
        """
        status_colunas = self.controlador.checar_colunas()
        for widget, status in zip(self.botoes_status, status_colunas):
            widget.definir_status(status)
        if all(stts == "ok" for stts in status_colunas):
            self.botao_gerar_modelo.setEnabled(True)
        else:
            self.botao_gerar_modelo.setEnabled(False)

    def icone_status_clicado(self, coluna: str, status: str):
        """Chama a função do controlador para identificar os problemas na coluna e mostra os resultados em uma popup.
        :param coluna: A coluna a ser verificada.
        :param status: O status da coluna ("ok", "faltando", "problemas", "nulos" ou "dominio").
        :return: Nada.
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
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro", parent=self)

    def botao_gerar_caderneta_clicado(self):
        """Chama as funções do controlador para gerar a caderneta e exportá-la.
        :return: Nada.
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
            mostrar_popup(f"ERRO: {exception}", tipo_msg="erro", parent=self)


class BotaoStatus(QtWidgets.QPushButton):
    def __init__(self, coluna: str, parent: JanelaPrincipalApp, status: str = "none"):
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
        """Define o ícone e a tooltip do botão. Conecta à função icone_status_clicado caso o status não seja "none" ou "ok".
        :param status: O status da coluna ("none", "ok", "faltando", "problemas", "nulos" ou "dominio")
        :return: Nada.
        """
        dic_botoes = {
            "none": {
                "icone": QtGui.QIcon("recursos_app/icones/circle.png"),
                "tooltip": "Carregue um arquivo"
            },
            "ok": {
                "icone": QtGui.QIcon("recursos_app/icones/ok.png"),
                "tooltip": "OK"
            },
            "missing_column": {
                "icone": QtGui.QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "Coluna não encontrada na tabela"
            },
            "wrong_dtype": {
                "icone": QtGui.QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna contém dados com\nformato errado"
            },
            "nan_not_allowed": {
                "icone": QtGui.QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "A coluna não permite nulos,\nmas existem células vazias"
            },
            "outside_domain": {
                "icone": QtGui.QIcon("recursos_app/icones/not_ok.png"),
                "tooltip": "Algumas células contêm valores\nfora da lista de valores permitidos"
            }
        }

        icone = QtGui.QIcon(dic_botoes[status]["icone"])
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


def mostrar_dialogo_arquivo(titulo: str, filtro: str, modo="abrir", parent: JanelaPrincipalApp = None):
    """Abre um diálogo de seleção/salvamento de arquivo.
    :param titulo: O título da janela.
    :param filtro: Filtros de tipo de arquivo (Ex: "Planilha do Excel (*.xlsx);;Planilha com macro do Excel (*.xlsm)")
    :param modo: "abrir" ou "salvar". Define se o diálogo será de abertura ou salvamento de arquivo.
    :param parent: A janela pai (Default = None).
    :return: Nada.
    """
    dialog = QtWidgets.QFileDialog(parent)
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
    """Troca o cursor do mouse por um cursor de espera.
    :param ativar: Default True. False para restaurar o cursor normal.
    :return: Nada.
    """
    if ativar:
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
    else:
        QtWidgets.QApplication.restoreOverrideCursor()


def mostrar_popup(mensagem: str, tipo_msg: str = "notificacao", parent: JanelaPrincipalApp = None):
    """Mostra uma popup com uma mensagem ao usuário.
    :param mensagem: A mensagem a ser exibida na popup.
    :param tipo_msg: "notificacao" ou "erro" (define o ícone da popup). O valor padrão é "notificacao".
    :param parent: A janela pai (Default = None).
    :return: Nada.
    """
    tipos_popup = {
        "notificacao": {"titulo": "Notificação", "icone": "recursos_app/icones/info.png"},
        "erro":        {"titulo": "Erro",        "icone": "recursos_app/icones/error.png"}
    }
    title = tipos_popup[tipo_msg]["titulo"]
    icon = QtGui.QIcon(tipos_popup[tipo_msg]["icone"])

    popup = QtWidgets.QMessageBox(parent)
    popup.setText(mensagem)
    popup.setWindowTitle(title)
    popup.setWindowIcon(icon)
    popup.exec()
