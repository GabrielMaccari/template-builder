import sys
import wx
import platform

from docx.opc import exceptions as docx_exceptions

from Controller import ControladorPrincipal, COLUNAS_TABELA_CADERNETA

TEMPLATE = f"recursos_app/modelos/template_estilos.docx"
OS = platform.platform()


class JanelaPrincipalApp(wx.Frame):
    def __init__(self):
        super(JanelaPrincipalApp, self).__init__(None)

        # Carrega o template de estilos da caderneta e instancia o controlador
        try:
            self.controlador = ControladorPrincipal(TEMPLATE)
        except docx_exceptions.PackageNotFoundError:
            mostrar_popup(f"Dependência não encontrada: {TEMPLATE}. Restaure o arquivo a partir do repositório e tente "
                          f"novamente.", "Erro", self)
            sys.exit()

        # Constrói a interface -----
        icone = wx.Icon()
        lxa = 20 if OS.startswith("Win") else 50
        icone.CopyFromBitmap(wx.Bitmap(wx.Image("recursos_app/icones/book.ico").Scale(lxa, lxa, wx.IMAGE_QUALITY_HIGH)))
        self.SetIcon(icone)
        self.SetTitle("Template Builder")

        painel = wx.Panel(self)

        layout_arquivo = wx.BoxSizer(wx.HORIZONTAL)

        self.rotulo_arquivo = wx.StaticText(painel, label="Selecione uma tabela com os dados dos pontos mapeados.")
        self.botao_abrir_arquivo = wx.Button(painel, label="Selecionar", size=(85, 30))

        layout_arquivo.Add(self.rotulo_arquivo, flag=wx.ALIGN_CENTRE_VERTICAL, proportion=1)
        layout_arquivo.Add(self.botao_abrir_arquivo, flag=wx.RIGHT)

        separador1 = wx.StaticLine(painel)

        layout_colunas = wx.FlexGridSizer(cols=4, vgap=7, hgap=20)
        layout_colunas.AddGrowableCol(0)
        layout_colunas.AddGrowableCol(2)
        layout_colunas.SetFlexibleDirection(wx.HORIZONTAL)

        # Cria os rótulos das colunas e botões de status em listas e adiciona-os ao layout em grade da seção central
        self.rotulos_colunas, self.botoes_status = [], []
        for coluna in COLUNAS_TABELA_CADERNETA:
            rotulo = wx.StaticText(painel, label=coluna)
            rotulo.SetMinSize((180,22))
            self.rotulos_colunas.append(rotulo)
            self.botoes_status.append(BotaoStatus(coluna, painel, self))

        tamanho_coluna_grid = int(len(self.rotulos_colunas)/2)
        if len(self.rotulos_colunas) % 2 != 0:
            tamanho_coluna_grid += 1
        i = 0
        while i < tamanho_coluna_grid:
            try:
                layout_colunas.Add(window=self.rotulos_colunas[i], flag=wx.EXPAND, proportion=1)
                layout_colunas.Add(window=self.botoes_status[i], flag=wx.CENTRE)
                layout_colunas.Add(window=self.rotulos_colunas[i+tamanho_coluna_grid], flag=wx.EXPAND, proportion=1)
                layout_colunas.Add(window=self.botoes_status[i+tamanho_coluna_grid], flag=wx.CENTRE)
            except IndexError:
                break
            i += 1
        
        layout_num_pontos = wx.BoxSizer(wx.HORIZONTAL)

        self.rotulo_num_pontos = wx.StaticText(painel, label="Número de pontos na tabela:   ")
        self.num_pontos = wx.StaticText(painel, label="-")

        layout_num_pontos.Add(self.rotulo_num_pontos)
        layout_num_pontos.Add(self.num_pontos)

        separador2 = wx.StaticLine(painel)

        self.checkbox_folha_rosto = wx.CheckBox(painel, label="Incluir folha de rosto no início da caderneta")
        self.checkbox_folha_rosto.SetValue(True)

        self.botao_gerar_caderneta = wx.Button(painel, label="Gerar caderneta")
        self.botao_gerar_caderneta.SetMinSize((530,47))
        self.botao_gerar_caderneta.Disable()

        layout_principal = wx.BoxSizer(wx.VERTICAL)
        layout_principal.Add(layout_arquivo, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)
        layout_principal.AddSpacer(5)
        layout_principal.Add(separador1, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)
        layout_principal.AddSpacer(5)
        layout_principal.Add(layout_colunas, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)
        layout_principal.AddSpacer(10)
        layout_principal.Add(layout_num_pontos, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)
        layout_principal.AddSpacer(5)
        layout_principal.Add(separador2, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)
        layout_principal.AddSpacer(10)
        layout_principal.Add(self.checkbox_folha_rosto, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=10)
        layout_principal.AddSpacer(10)
        layout_principal.Add(self.botao_gerar_caderneta, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        layout_principal.AddSpacer(46 if OS.startswith("Win") else 31)

        painel.SetSizerAndFit(layout_principal)

        self.botao_abrir_arquivo.Bind(wx.EVT_BUTTON, self.botao_abrir_arquivo_clicado)
        self.botao_gerar_caderneta.Bind(wx.EVT_BUTTON, self.botao_gerar_caderneta_clicado)

        self.SetSize(painel.Size)
        print(layout_colunas.GetColWidths())

    def botao_abrir_arquivo_clicado(self, evento: wx.Event):
        """Chama a função para abrir a tabela no controlador. Caso um arquivo seja aberto com sucesso, atualiza o
        rotulo_arquivo com o nome do arquivo selecionado e chama o método checar_colunas para atualizar os ícones de
        status.
        :return: Nada.
        """
        try:
            arquivo_aberto, num_pontos = False, "-"
            caminho = selecionar_arquivo(parent=self)
            if caminho is not None:
                arquivo_aberto, num_pontos = self.controlador.abrir_tabela(caminho)
            if arquivo_aberto:
                partes_caminho = caminho.split("\\" if OS.startswith("Win") else "/")
                self.rotulo_arquivo.SetLabel(partes_caminho[-1])
                self.num_pontos.SetLabel(str(num_pontos) if num_pontos > 0 else "-")
                self.checar_colunas()

                if "nan" in self.controlador.df.columns:
                    mostrar_popup("Atenção! Existem colunas com nomes inválidos na tabela que podem causar erros ou "
                                  "anomalias no funcionamento da ferramenta. Verifique se as fórmulas presentes nas "
                                  "células de cabeçalho das colunas de estruturas (colunas S a AG) não foram "
                                  "comprometidas. Isso geralmente ocorre ao recortar e colar células na aba de Listas "
                                  "ao preencher as estruturas.", parent=self)

        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", "Erro", self)

    def checar_colunas(self):
        """Chama a função do controlador para checar se cada coluna está no formato especificado. Atualiza os botões de
        status conforme o resultado. Se todas as colunas estiverem OK, habilita o botão para gerar o template.
        :return: Nada.
        """
        status_colunas = self.controlador.checar_colunas()
        for botao, status in zip(self.botoes_status, status_colunas):
            botao.definir_status(status)
        if all(stts == "ok" for stts in status_colunas):
            self.botao_gerar_caderneta.Enable()
        else:
            self.botao_gerar_caderneta.Disable()

    def botao_status_clicado(self, coluna: str, status: str):
        """Chama a função do controlador para identificar os problemas na coluna e mostra os resultados em uma popup.
        :param coluna: A coluna a ser verificada.
        :param status: O status da coluna ("ok", "faltando", "problemas", "nulos" ou "dominio").
        :return: Nada.
        """
        try:
            localizar_problemas = {
                "missing_column": lambda missing_column: [],
                "wrong_dtype": self.controlador.localizar_problemas_formato,
                "nan_not_allowed": self.controlador.localizar_celulas_vazias,
                "outside_domain": self.controlador.localizar_problemas_dominio
            }

            if status not in localizar_problemas.keys():
                raise Exception(f"O status informado não foi reconhecido: {status}")

            indices_problemas = localizar_problemas[status](coluna)

            msg = self.controlador.montar_msg_problemas(status, coluna, indices_problemas)
            mostrar_popup(msg, parent=self)

        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", "Erro", self)

    def botao_gerar_caderneta_clicado(self, evento: wx.Event):
        """Chama as funções do controlador para gerar a caderneta e exportá-la.
        :return: Nada.
        """
        try:
            cursor_espera = wx.BusyCursor()
            montar_folha_rosto = self.checkbox_folha_rosto.GetValue()
            self.controlador.gerar_caderneta(montar_folha_rosto)
            del cursor_espera
            caminho = selecionar_arquivo(self, "salvar", "Documento do Word (*.docx)|*.docx")
            if caminho is not None:
                self.controlador.salvar_caderneta(caminho)
                mostrar_popup("Caderneta criada com sucesso!", parent=self)
        except Exception as exception:
            mostrar_popup(f"ERRO: {exception}", "Erro", self)


class BotaoStatus(wx.Button):
    def __init__(self, coluna: str, container: wx.Panel, janela: JanelaPrincipalApp, status: str = "none"):
        super().__init__(container, wx.ID_ANY, style=wx.BU_NOTEXT|wx.BORDER_NONE, size=(22, 22))

        self.coluna = coluna
        self.parent = janela
        self.status = status

        self.definir_status(status)

    def definir_status(self, status: str = "none"):
        dic_botoes = {
            "none": {
                "icone": wx.Image("recursos_app/icones/circle_gray.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": None
            },
            "ok": {
                "icone": wx.Image("recursos_app/icones/ok.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": "OK"
            },
            "missing_column": {
                "icone": wx.Image("recursos_app/icones/not_ok.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": "Coluna não encontrada na tabela"
            },
            "wrong_dtype": {
                "icone": wx.Image("recursos_app/icones/not_ok.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": "A coluna contém dados com\nformato errado"
            },
            "nan_not_allowed": {
                "icone": wx.Image("recursos_app/icones/not_ok.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": "A coluna não permite nulos,\nmas existem células vazias"
            },
            "outside_domain": {
                "icone": wx.Image("recursos_app/icones/not_ok.png", wx.BITMAP_TYPE_ANY).Scale(20, 20, wx.IMAGE_QUALITY_HIGH),
                "tooltip": "Algumas células contêm valores\nfora da lista de valores permitidos"
            }
        }

        icone = wx.Bitmap(dic_botoes[status]["icone"])
        tooltip = dic_botoes[status]["tooltip"]

        self.status = status
        self.SetBitmap(icone)
        self.SetToolTip(tooltip)

        try:
            self.Unbind(wx.EVT_BUTTON)
        except Exception as e:
            print(e.__class__)

        if status == "none":
            self.Disable()
        else:
            self.Enable()

        if status not in ["none", "ok"]:
            self.Bind(wx.EVT_BUTTON, lambda x: self.parent.botao_status_clicado(self.coluna, self.status))


def mostrar_popup(mensagem: str, titulo: str = "Notificação", parent: JanelaPrincipalApp = None):
    """Mostra uma popup com uma mensagem ao usuário.
    :param mensagem: A mensagem a ser exibida na popup.
    :param titulo: "Notificação" ou "Erro" (define o ícone da popup). O valor padrão é "Notificação".
    :param parent: A janela pai (Default = None).
    :return: Nada.
    """
    estilo = wx.OK|wx.CENTRE|wx.ICON_ERROR if titulo == "Erro" else wx.OK|wx.CENTRE|wx.ICON_INFORMATION
    popup = wx.MessageDialog(parent, mensagem, titulo, style=estilo)
    popup.ShowModal()
    popup.Destroy()


def selecionar_arquivo(parent: JanelaPrincipalApp = None, modo: str = "abrir",
                       filtro: str = "Pastas de Trabalho do Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm") -> str|None:
    """Abre um diálogo de seleção/salvamento de arquivo.
    :param parent: A janela pai (Default = None).
    :param modo: "open" ou "save". Define se o diálogo será de abertura ou salvamento de arquivo.
    :param filtro: Filtros de tipo de arquivo (Ex: "Planilha do Excel (\*.xlsx)|CSV (\*.csv)")
    :return: Nada.
    """
    titulo = "Selecionar arquivo" if modo == "abrir" else "Salvar arquivo"
    estilo = wx.FD_OPEN|wx.FD_FILE_MUST_EXIST if modo == "abrir" else wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT

    with wx.FileDialog(parent, titulo, wildcard=filtro, style=estilo) as dialogo_arquivo:
        if dialogo_arquivo.ShowModal() == wx.ID_CANCEL:
            return None
        return dialogo_arquivo.GetPath()
