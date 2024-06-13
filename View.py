# -*- coding: utf-8 -*-
"""
@author: Gabriel Maccari
"""

from PyQt6.QtWidgets import *
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt

from Model import COLUNAS_TABELA_CADERNETA


class Interface(QMainWindow):
    def __init__(self):
        super(Interface, self).__init__(None)

        # Função a ser executada quando o usuário clica nos ícones de status das colunas (é definida no controlador)
        self.funcao_botoes_status = None

        # Constrói a interface -----
        self.setWindowTitle('Template Builder')
        self.setWindowIcon(QIcon('config/icones/book.png'))

        # Rótulo do nome do arquivo
        self.rotulo_arquivo = QLabel("Selecione um arquivo .xlsx com os dados dos pontos mapeados.")

        # Botão para selecionar a tabela
        self.botao_abrir_arquivo = QPushButton("Selecionar")
        self.botao_abrir_arquivo.setMaximumWidth(75)

        # Layout que organiza a porção superior da interface (acima da linha)
        layout_superior = QHBoxLayout()

        layout_superior.addWidget(self.rotulo_arquivo)
        layout_superior.addWidget(self.botao_abrir_arquivo)

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

        # Cria os rótulos das colunas e botões de status em listas e adiciona-os ao layout em grade da seção central
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

        # Checkbox para marcar se gera ou não as folhas de título dos semestres
        self.checkbox_folhas_semestre = QCheckBox("Incluir folhas de título para cada semestre/disciplina")
        self.checkbox_folhas_semestre.setToolTip("Gera páginas com o título da disciplina antes do primeiro ponto de "
                                                 "cada semestre")
        self.checkbox_folhas_semestre.setChecked(True)

        # Checkbox para marcar se o usuário deseja continuar uma caderneta já existente
        self.checkbox_continuar_caderneta = QCheckBox("Continuar caderneta existente")
        self.checkbox_continuar_caderneta.setToolTip("Utilize esta opção para adicionar novos pontos a uma caderneta\n"
                                                     "pré-existente (Ex: adicionar pontos do Map2 à caderneta do Map1)")

        # Layout que organiza os widgets de seleção do ponto de início
        layout_selecao_ponto_inicio = QHBoxLayout()

        # Rótulo da seleção de ponto inicial
        self.rotulo_ponto_inicio = QLabel("Iniciar a partir do ponto: ")
        self.rotulo_ponto_inicio.setEnabled(False)

        # Caixa de seleção do ponto inicial
        self.combobox_ponto_inicio = QComboBox()
        self.combobox_ponto_inicio.setStyleSheet("combobox-popup: 0;")
        self.combobox_ponto_inicio.setMinimumWidth(100)
        self.combobox_ponto_inicio.setEnabled(False)

        layout_selecao_ponto_inicio.addWidget(self.rotulo_ponto_inicio)
        layout_selecao_ponto_inicio.addWidget(self.combobox_ponto_inicio)
        layout_selecao_ponto_inicio.addStretch(1)

        # Botão para gerar e exportar o template da caderneta
        self.botao_gerar_nova_caderneta = QPushButton("Gerar caderneta")
        self.botao_gerar_nova_caderneta.setMinimumHeight(35)
        self.botao_gerar_nova_caderneta.setEnabled(False)

        # Layout mestre (aninha os widgets e demais layouts)
        layout_principal = QVBoxLayout()
        layout_principal.setSpacing(5)

        layout_principal.addLayout(layout_superior)
        layout_principal.addWidget(separador)
        layout_principal.addLayout(layout_central)
        layout_principal.addWidget(separador2)
        layout_principal.addWidget(self.checkbox_folha_rosto)
        layout_principal.addWidget(self.checkbox_folhas_semestre)
        layout_principal.addWidget(self.checkbox_continuar_caderneta)
        layout_principal.addLayout(layout_selecao_ponto_inicio)
        layout_principal.addWidget(self.botao_gerar_nova_caderneta)

        # Painel contendo o layout
        container = QWidget(self)
        container.setLayout(layout_principal)
        self.setCentralWidget(container)


class BotaoStatus(QPushButton):
    def __init__(self, coluna: str, parent: Interface, status: str = "none"):
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
                "icone": QIcon("config/icones/circle.png"),
                "tooltip": "Carregue um arquivo"
            },
            "ok": {
                "icone": QIcon("config/icones/ok.png"),
                "tooltip": "OK"
            },
            "coluna_faltando": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "Coluna não encontrada na tabela"
            },
            "fora_de_formato": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "A coluna contém dados com\nformato errado"
            },
            "celulas_vazias": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "A coluna não permite nulos,\nmas existem células vazias"
            },
            "valores_nao_permitidos": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "Algumas células contêm valores\nfora da lista de valores permitidos"
            },
            "fora_do_intervalo": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "Algumas células contêm valores\nnuméricos fora do intervalo permitido"
            },
            "valores_repetidos": {
                "icone": QIcon("config/icones/not_ok.png"),
                "tooltip": "Existem valores repetidos"
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
            self.clicked.connect(lambda: self.parent.funcao_botoes_status(self.coluna, self.status))

        self.setEnabled(status != "none")
