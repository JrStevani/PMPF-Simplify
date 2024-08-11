from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5 import QtCore, QtWidgets, QtGui
from os import path, makedirs
from datetime import datetime
import pandas as pd
import numpy as np
import traceback
import sys


def ordenar_grupo(tabela):
    return tabela.groupby(by='CODG_EAN').apply(lambda x: x.sort_values(by='VALR_UNIDADE_CALC', ascending=True))


def diferca_de_pmpf(antigo, novo):
    diferenca_percentual = ((novo - antigo) / abs(antigo)) * 100
    return str(f'{diferenca_percentual:.2f}')


def calcular_intervalo_aceitacao(valores):
    if not valores:
        return [0, 0]

    desvio_padrao = np.std(valores)
    media_valores = np.mean(valores)

    inferior = float(media_valores - (3 * desvio_padrao))
    superior = float(media_valores + (3 * desvio_padrao))

    return inferior, superior


def salvar_dataframe_para_excel(nome_arquivo, nome_aba, dataframe, w=None):
    if w is None:
        w = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
    dataframe.to_excel(w, sheet_name=nome_aba, index=False)
    return w


def encontrar_valor_mais_proximo(lista_de_valores, media_ponderavel):
    menor_diferenca = float('inf')
    valor_da_lista = None

    for x in lista_de_valores:
        diferenca = abs(x - media_ponderavel)
        if diferenca < menor_diferenca:
            menor_diferenca = diferenca
            valor_da_lista = x
    return valor_da_lista


def calcular_quantidade_valor_unitario(linha_a, list_cod_gtin, valor_min, fitro_valor_minimo):
    cod_gtin = str(linha_a['CODG_EAN'])
    list_cod_gtin = list(list_cod_gtin[cod_gtin])
    setor = float(list_cod_gtin[int(linha_a['ÍNDICES_FORNECEDOR']) + 2])
    qtde_comercial = float(linha_a['QTDE_COMERCIAL'])
    valor_produto = float(linha_a['VALR_PRODUTO'])
    pmpf_ = float(list_cod_gtin[2])
    lista_de_caixas = list(str(list_cod_gtin[-1]).replace('.', ',').split(','))

    try:
        if len(lista_de_caixas) == 1:
            QUANT_PONDERAVEIS = int(lista_de_caixas[0])
            QTDE_COMERCIAL_CALC = QUANT_PONDERAVEIS * qtde_comercial
            VALR_UND_CALC = valor_produto / QTDE_COMERCIAL_CALC
        else:
            if setor != 0:
                pmpf_ = setor
            lista_de_valores = []
            valor_min = float(valor_min) / 100

            valor_min = pmpf_ * valor_min
            for item in lista_de_caixas:
                valor_unitario = valor_produto / (int(item) * qtde_comercial)
                if valor_unitario >= valor_min and valor_unitario >= fitro_valor_minimo:
                    lista_de_valores.append(valor_unitario)
            if len(lista_de_valores) == 0:
                QUANT_PONDERAVEIS = int(lista_de_caixas[0])
                QTDE_COMERCIAL_CALC = QUANT_PONDERAVEIS * qtde_comercial
                VALR_UND_CALC = valor_produto / QTDE_COMERCIAL_CALC
            elif len(lista_de_valores) == 1:
                menor_valor = lista_de_valores[0]
                QUANT_PONDERAVEIS = valor_produto / (menor_valor * qtde_comercial)
                QTDE_COMERCIAL_CALC = QUANT_PONDERAVEIS * qtde_comercial
                VALR_UND_CALC = valor_produto / QTDE_COMERCIAL_CALC
            else:
                menor_valor = encontrar_valor_mais_proximo(lista_de_valores, pmpf_)
                QUANT_PONDERAVEIS = valor_produto / (menor_valor * qtde_comercial)
                QTDE_COMERCIAL_CALC = QUANT_PONDERAVEIS * qtde_comercial
                VALR_UND_CALC = valor_produto / QTDE_COMERCIAL_CALC
        return pd.Series({'VALR_UNIDADE_CALC': round(VALR_UND_CALC, 2), 'QTDE_COMERCIAL_CALC': QTDE_COMERCIAL_CALC,
                          'QUANT_PONDERAVEIS': QUANT_PONDERAVEIS})

    except Exception as e:
        print(e)

        return pd.Series({'VALR_UNIDADE_CALC': 0, 'QTDE_COMERCIAL_CALC': 0, 'QUANT_PONDERAVEIS': 0})


def corrigeir_gtin(r):
    return str(r['CODG_EAN']).split('.')[0]


def formatar_para_float(valor):
    valor = str(valor).replace("'", "").replace('"', '')
    if ',' in valor:
        centavos = str(valor).split(',')[1]
        reais = str(valor).split(',')[0].replace(',', '').replace('.', '')
        return round(float(f'{reais}.{centavos}'), 2)
    elif '.' in str(valor):
        centavos = str(valor).split('.')[1]
        reais = str(valor).split('.')[0].replace(',', '').replace('.', '')
        return round(float(f'{reais}.{centavos}'), 2)
    else:
        return round(float(valor), 2)


def remover_colunas(df_tabela):
    df_tabela.rename(columns=lambda x: x.upper().replace('\n', '').replace('  ', ' ').strip(), inplace=True)
    return df_tabela[['NUMR_INSC_ESTADUAL_EMISSOR', 'NOME_FANTASIA_EMISSOR',
                      'DESC_PRODUTO', 'CODG_EAN', 'VALR_UNIDADE_COMERCIAL',
                      'UNID_COMERCIAL', 'QTDE_COMERCIAL', 'VALR_PRODUTO']]


def hora():
    agora = datetime.now()
    return agora.strftime("%H:%M:%S")


def registrar_erro(nome_arquivo, texto):
    try:
        with open(nome_arquivo, 'r', encoding='utf-8') as file_a:
            conteudo_atual = file_a.read()
        with open(nome_arquivo, 'a', encoding='utf-8') as file_b:
            if conteudo_atual:
                file_b.write('\n' + '=-' * 35 + '\n')
            file_b.write(texto)

    except FileNotFoundError:
        with open(nome_arquivo, 'w', encoding='utf-8') as file_c:
            file_c.write(texto)


def configurar_ambiente():
    if not path.exists(f'.\\confg'):
        makedirs(f'.\\confg')

    if not path.exists('.\\confg\\Erros.txt'):
        with open('.\\confg\\Erros.txt', 'w', encoding='utf-8') as file_d:
            file_d.write('=' * 15 + ' REGISTRO DE ERROS ' + '=' * 15 + '\n')


class Ui_Form(object):

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1040, 495)
        icon = QtGui.QIcon(r".\\confg\\icon.ico")
        Form.setWindowIcon(icon)

        self.tabWidget = QtWidgets.QTabWidget(Form)
        self.tabWidget.setEnabled(True)
        self.tabWidget.setGeometry(QtCore.QRect(10, 10, 1020, 465))
        self.tabWidget.setObjectName("tabWidget")

        self.tab_bebidas = QtWidgets.QWidget()
        self.tab_bebidas.setObjectName("tab_bebidas")

        self.checkBox_bebidas = QtWidgets.QCheckBox(self.tab_bebidas)
        self.checkBox_bebidas.setGeometry(QtCore.QRect(750, 120, 181, 41))
        self.checkBox_bebidas.setChecked(True)
        self.checkBox_bebidas.setObjectName("checkBox_bebidas")

        self.checkBox2_bebidas = QtWidgets.QCheckBox(self.tab_bebidas)
        self.checkBox2_bebidas.setGeometry(QtCore.QRect(750, 160, 181, 41))
        self.checkBox2_bebidas.setChecked(True)
        self.checkBox2_bebidas.setObjectName("checkBox_bebidas")

        self.checkBox3_bebidas = QtWidgets.QCheckBox(self.tab_bebidas)
        self.checkBox3_bebidas.setGeometry(QtCore.QRect(750, 200, 181, 41))
        self.checkBox3_bebidas.setChecked(True)
        self.checkBox3_bebidas.setObjectName("checkBox_bebidas")

        self.pushButton_2_bebidas = QtWidgets.QPushButton(self.tab_bebidas)
        self.pushButton_2_bebidas.setGeometry(QtCore.QRect(750, 20, 171, 31))
        self.pushButton_2_bebidas.setObjectName("pushButton_2_bebidas")
        self.pushButton_2_bebidas.clicked.connect(self.open_file_dialog_bebidas)

        self.pushButton_3_bebidas = QtWidgets.QPushButton(self.tab_bebidas)
        self.pushButton_3_bebidas.setGeometry(QtCore.QRect(750, 70, 171, 31))
        self.pushButton_3_bebidas.setObjectName("pushButton_3_bebidas")
        self.pushButton_3_bebidas.clicked.connect(self.iniciar_bebidas)

        self.line_bebidas = QtWidgets.QFrame(self.tab_bebidas)
        self.line_bebidas.setGeometry(QtCore.QRect(700, 10, 20, 291))
        self.line_bebidas.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_bebidas.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_bebidas.setObjectName("line_bebidas")

        self.radioButton_bebidas = QtWidgets.QRadioButton(self.tab_bebidas)
        self.radioButton_bebidas.setGeometry(QtCore.QRect(20, 60, 211, 51))
        self.radioButton_bebidas.setObjectName("radioButton_bebidas")

        self.radioButton_2_bebidas = QtWidgets.QRadioButton(self.tab_bebidas)
        self.radioButton_2_bebidas.setGeometry(QtCore.QRect(240, 60, 211, 51))
        self.radioButton_2_bebidas.setObjectName("radioButton_2_bebidas")

        self.label_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_bebidas.setGeometry(QtCore.QRect(20, 10, 170, 41))
        self.label_bebidas.setObjectName("label_bebidas")

        self.txt_file_path_bebidas = QtWidgets.QLineEdit(self.tab_bebidas)
        self.txt_file_path_bebidas.setGeometry(QtCore.QRect(200, 15, 420, 30))
        self.txt_file_path_bebidas.mouseDoubleClickEvent = lambda event: self.open_file_dialog_bebidas()
        open_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.CTRL + QtCore.Qt.Key_O), self.txt_file_path_bebidas)
        open_shortcut.activated.connect(self.open_file_dialog_bebidas)

        self.label_2_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_2_bebidas.setGeometry(QtCore.QRect(26, 115, 141, 61))
        self.label_2_bebidas.setObjectName("label_2_bebidas")

        self.label_2_1_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_2_1_bebidas.setGeometry(QtCore.QRect(26, 165, 141, 61))
        self.label_2_1_bebidas.setObjectName("label_2_1_bebidas")

        self.txt_file_path_2_bebidas = QtWidgets.QLineEdit(self.tab_bebidas)
        self.txt_file_path_2_bebidas.setGeometry(QtCore.QRect(150, 130, 120, 30))

        self.txt_file_path_2_1_bebidas = QtWidgets.QLineEdit(self.tab_bebidas)
        self.txt_file_path_2_1_bebidas.setGeometry(QtCore.QRect(180, 180, 120, 30))

        self.label_3_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_3_bebidas.setGeometry(QtCore.QRect(380, 260, 100, 31))
        self.label_3_bebidas.setObjectName("label_3_bebidas")

        self.label_4_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_4_bebidas.setGeometry(QtCore.QRect(490, 260, 100, 31))
        self.label_4_bebidas.setObjectName("label_4_bebidas")

        self.label_5_bebidas = QtWidgets.QLabel(self.tab_bebidas)
        self.label_5_bebidas.setGeometry(QtCore.QRect(30, 260, 290, 31))
        self.label_5_bebidas.setObjectName("label_5_bebidas")

        self.progressBar_2_bebidas = QtWidgets.QProgressBar(self.tab_bebidas)
        self.progressBar_2_bebidas.setGeometry(QtCore.QRect(20, 370, 1000, 21))
        self.progressBar_2_bebidas.setProperty("value", 0)
        self.progressBar_2_bebidas.setObjectName("progressBar_2_bebidas")

        # Interface de medicamentos

        self.tabWidget.addTab(self.tab_bebidas, "")
        self.tab_medicamentos = QtWidgets.QWidget()
        self.tab_medicamentos.setObjectName("tab_medicamentos")
        self.tabWidget.addTab(self.tab_medicamentos, "")

        self.checkBox_medicamentos = QtWidgets.QCheckBox(self.tab_medicamentos)
        self.checkBox_medicamentos.setGeometry(QtCore.QRect(750, 120, 181, 41))
        self.checkBox_medicamentos.setChecked(True)
        self.checkBox_medicamentos.setObjectName("checkBox_medicamentos")

        self.checkBox2_medicamentos = QtWidgets.QCheckBox(self.tab_medicamentos)
        self.checkBox2_medicamentos.setGeometry(QtCore.QRect(750, 160, 181, 41))
        self.checkBox2_medicamentos.setChecked(True)
        self.checkBox2_medicamentos.setObjectName("checkBox_medicamentos")

        self.checkBox3_medicamentos = QtWidgets.QCheckBox(self.tab_medicamentos)
        self.checkBox3_medicamentos.setGeometry(QtCore.QRect(750, 200, 181, 41))
        self.checkBox3_medicamentos.setChecked(True)
        self.checkBox3_medicamentos.setObjectName("checkBox_medicamentos")

        self.pushButton_2_medicamentos = QtWidgets.QPushButton(self.tab_medicamentos)
        self.pushButton_2_medicamentos.setGeometry(QtCore.QRect(750, 20, 171, 31))
        self.pushButton_2_medicamentos.setObjectName("pushButton_2_medicamentos")
        self.pushButton_2_medicamentos.clicked.connect(self.open_file_dialog_medicamentos)

        self.pushButton_3_medicamentos = QtWidgets.QPushButton(self.tab_medicamentos)
        self.pushButton_3_medicamentos.setGeometry(QtCore.QRect(750, 70, 171, 31))
        self.pushButton_3_medicamentos.setObjectName("pushButton_3_medicamentos")
        self.pushButton_3_medicamentos.clicked.connect(self.iniciar_medicamentos)

        self.line_medicamentos = QtWidgets.QFrame(self.tab_medicamentos)
        self.line_medicamentos.setGeometry(QtCore.QRect(700, 10, 20, 291))
        self.line_medicamentos.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_medicamentos.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_medicamentos.setObjectName("line_medicamentos")

        self.radioButton_medicamentos = QtWidgets.QRadioButton(self.tab_medicamentos)
        self.radioButton_medicamentos.setGeometry(QtCore.QRect(20, 60, 211, 51))
        self.radioButton_medicamentos.setObjectName("radioButton_medicamentos")

        self.radioButton_2_medicamentos = QtWidgets.QRadioButton(self.tab_medicamentos)
        self.radioButton_2_medicamentos.setGeometry(QtCore.QRect(240, 60, 211, 51))
        self.radioButton_2_medicamentos.setObjectName("radioButton_2_medicamentos")

        self.label_medicamentos = QtWidgets.QLabel(self.tab_medicamentos)
        self.label_medicamentos.setGeometry(QtCore.QRect(20, 10, 170, 41))
        self.label_medicamentos.setObjectName("label_medicamentos")

        self.txt_file_path_medicamentos = QtWidgets.QLineEdit(self.tab_medicamentos)
        self.txt_file_path_medicamentos.setGeometry(QtCore.QRect(200, 15, 420, 30))
        self.txt_file_path_medicamentos.mouseDoubleClickEvent = lambda event: self.open_file_dialog_medicamentos()
        open_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.CTRL + QtCore.Qt.Key_O),
                                            self.txt_file_path_medicamentos)
        open_shortcut.activated.connect(self.open_file_dialog_medicamentos)

        self.label_2_medicamentos = QtWidgets.QLabel(self.tab_medicamentos)
        self.label_2_medicamentos.setGeometry(QtCore.QRect(26, 115, 141, 61))
        self.label_2_medicamentos.setObjectName("label_2_medicamentos")

        self.txt_file_path_2_medicamentos = QtWidgets.QLineEdit(self.tab_medicamentos)
        self.txt_file_path_2_medicamentos.setGeometry(QtCore.QRect(150, 130, 120, 30))

        self.label_3_medicamentos = QtWidgets.QLabel(self.tab_medicamentos)
        self.label_3_medicamentos.setGeometry(QtCore.QRect(380, 260, 100, 31))
        self.label_3_medicamentos.setObjectName("label_3_medicamentos")

        self.label_4_medicamentos = QtWidgets.QLabel(self.tab_medicamentos)
        self.label_4_medicamentos.setGeometry(QtCore.QRect(490, 260, 100, 31))
        self.label_4_medicamentos.setObjectName("label_4_medicamentos")

        self.label_5_medicamentos = QtWidgets.QLabel(self.tab_medicamentos)
        self.label_5_medicamentos.setGeometry(QtCore.QRect(30, 260, 290, 31))
        self.label_5_medicamentos.setObjectName("label_5_medicamentos")

        self.progressBar_2_medicamentos = QtWidgets.QProgressBar(self.tab_medicamentos)
        self.progressBar_2_medicamentos.setGeometry(QtCore.QRect(20, 370, 1000, 21))
        self.progressBar_2_medicamentos.setProperty("value", 0)
        self.progressBar_2_medicamentos.setObjectName("progressBar_2_medicamentos")


        self.retranslateUi(Form)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(Form)

        self.msgBox = QMessageBox()
        self.msgBox.setIcon(QMessageBox.Critical)
        self.msgBox.setWindowTitle("Erro")

        Form.setStyleSheet("""
            /* Estilo da QTabWidget */
            QTabWidget {
                background-color: #02E3BF;
                color: black;
                font-family: "Arial Narrow Bold";
            }
            
            QTabBar::tab {
                padding: 10px;
                font-family: "Arial Narrow Bold";
            }
            
            QTabBar::tab:hover {
                background-color: #B8E3BC;
                border-radius: 5px;
            }
            
            /* Estilo da janela */
            QWidget {
                background-color: #04E063;
                color: black;
                font-size: 11pt;
                font-family: serif;
            }

            /* Estilo dos botões */
            QPushButton {
                background-color: #53E091;
                border: 1px solid #333;
                color: black;
                font-size: 10pt;
                font-family: serif;
                padding: 5px 10px;
            }

            QPushButton:hover {
                background-color: #B8E3BC;
            }

            QPushButton:pressed {
                background-color: #42E3BF;
            }

            /* Estilo dos botões de rádio */
            QRadioButton {
                color: black;
                font-size: 10pt;
                font-family: serif;
            }

            QLineEdit {
                font-size: 10pt;
                background-color: white;
            }
        """)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "PMPF Simplify 9.1.0"))
        self.checkBox_bebidas.setText(_translate("Form", "Aplicar filtro"))
        self.checkBox_medicamentos.setText(_translate("Form", "Aplicar filtro"))
        self.checkBox2_bebidas.setText(_translate("Form", "Remover duplicatas"))
        self.checkBox2_medicamentos.setText(_translate("Form", "Remover duplicatas"))
        self.checkBox3_bebidas.setText(_translate("Form", "Filtro I.E.E"))
        self.checkBox3_medicamentos.setText(_translate("Form", "Filtro I.E.E"))

        self.pushButton_2_bebidas.setText(_translate("Form", "Selecionar Arquivo"))
        self.pushButton_3_bebidas.setText(_translate("Form", "Iniciar"))
        self.pushButton_2_medicamentos.setText(_translate("Form", "Selecionar Arquivo"))
        self.pushButton_3_medicamentos.setText(_translate("Form", "Iniciar"))

        self.radioButton_bebidas.setText(_translate("Form", "Gerar tabela padronizada"))
        self.radioButton_2_bebidas.setText(_translate("Form", "Calcular PMPF"))
        self.label_bebidas.setText(_translate("Form", "Selecione um arquivo"))
        self.label_2_bebidas.setText(_translate("Form", "Valor percentil"))
        self.label_2_1_bebidas.setText(_translate("Form", "Excluir menor que"))
        self.label_3_bebidas.setText(_translate("Form", ""))
        self.label_4_bebidas.setText(_translate("Form", ""))
        self.label_5_bebidas.setText(_translate("Form", ""))
        self.radioButton_medicamentos.setText(_translate("Form", "Gerar tabela padronizada"))
        self.radioButton_2_medicamentos.setText(_translate("Form", "Calcular PMPF"))
        self.label_medicamentos.setText(_translate("Form", "Selecione um arquivo"))
        self.label_2_medicamentos.setText(_translate("Form", "Valor percentil"))
        self.label_3_medicamentos.setText(_translate("Form", ""))
        self.label_4_medicamentos.setText(_translate("Form", ""))
        self.label_5_medicamentos.setText(_translate("Form", ""))

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_bebidas), _translate("Form", "Bebidas"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_medicamentos), _translate("Form", "Medicamentos"))

    def open_file_dialog_bebidas(self):
        options = QFileDialog.Options()

        file_names, _ = QFileDialog.getOpenFileNames(None, "Selecionar Arquivos", "", "Arquivos xlsx (*.xlsx)",
                                                     options=options)
        if file_names:
            file_names_str = "; ".join(file_names)
            self.txt_file_path_bebidas.setText(file_names_str)

    def open_file_dialog_medicamentos(self):
        options = QFileDialog.Options()
        file_names, _ = QFileDialog.getOpenFileNames(None, "Selecionar Arquivos", "", "Arquivos xlsx (*.xlsx)",
                                                     options=options)
        if file_names:
            file_names_str = "; ".join(file_names)
            self.txt_file_path_medicamentos.setText(file_names_str)

    def iniciar_bebidas(self):
        caminho_dos_arquivos = list(str(self.txt_file_path_bebidas.text()).split(';'))
        valor_minimo = self.txt_file_path_2_bebidas.text()
        fitro_valor_minimo = self.txt_file_path_2_1_bebidas.text()
        if fitro_valor_minimo:
            fitro_valor_minimo = float(str(fitro_valor_minimo).replace(',', '.'))
        else:
            fitro_valor_minimo = 0
        tarefa = self.radioButton_bebidas.isChecked()
        usar_filtro = self.checkBox_bebidas.isChecked()
        remover_duplicatas = self.checkBox2_bebidas.isChecked()
        filtro_iee = self.checkBox3_bebidas.isChecked()

        for caminho_arquivo in caminho_dos_arquivos:
            try:
                if (not tarefa and caminho_arquivo) or (caminho_arquivo and valor_minimo and valor_minimo.isnumeric()):
                    self.progressBar_2_bebidas.setValue(0)
                    QtWidgets.QApplication.processEvents()
                    self.label_5_bebidas.setText("Carregando...")
                    self.label_3_bebidas.setText(f"| {hora()} ")

                    if tarefa:
                        self.label_5_bebidas.setText("Abrindo tabela.")

                        file_paths = {
                            "Filtro": ".\\confg\\Filtro.xlsx",
                            "Gtin": ".\\confg\\Gtin.xlsx",
                            "Fornecedores": ".\\confg\\Fornecedores.xlsx"
                        }
                        dataframes = {}

                        for key, file in file_paths.items():
                            if path.exists(file):
                                dataframes[key] = pd.read_excel(io=file, sheet_name=0)
                            else:
                                dataframes[key] = pd.DataFrame()

                        df_linhas_separadas = pd.DataFrame()
                        df_de_erros = pd.DataFrame()

                        df_filtro = dataframes.get("Filtro", pd.DataFrame())
                        df_gtin = dataframes.get("Gtin", pd.DataFrame())
                        df_fornecedores = dataframes.get("Fornecedores", pd.DataFrame())
                        del dataframes
                        if str(caminho_arquivo.split('.')[-1]) == 'xlsx':
                            df_tabela_inteira = pd.read_excel(io=caminho_arquivo, sheet_name=0)
                            arquivo_excel = f'{str(caminho_arquivo).replace(".xlsx", " - nova.xlsx")}'
                            self.progressBar_2_bebidas.setValue(10)

                            # 1° Removendo duplicatas
                            if remover_duplicatas:
                                self.label_5_bebidas.setText("Removendo duplicatas.")
                                df_tabela_inteira = df_tabela_inteira.drop_duplicates()
                                df_tabela_inteira = remover_colunas(df_tabela_inteira)
                                self.progressBar_2_bebidas.setValue(20)

                            # 2° Formatando colunas
                            self.label_5_bebidas.setText('Formatando colunas.')
                            df_tabela_inteira['VALR_PRODUTO'] = df_tabela_inteira['VALR_PRODUTO'].apply(
                                formatar_para_float)
                            df_tabela_inteira['VALR_UNIDADE_COMERCIAL'] = df_tabela_inteira[
                                'VALR_UNIDADE_COMERCIAL'].apply(formatar_para_float)

                            if fitro_valor_minimo != 0:
                                # removendo VALR_PRODUTO menores que 0,99
                                linhas_invalidas = df_tabela_inteira[
                                    df_tabela_inteira['VALR_UNIDADE_COMERCIAL'] < fitro_valor_minimo].copy()

                                df_tabela_inteira = df_tabela_inteira[
                                    ~df_tabela_inteira.index.isin(linhas_invalidas.index)]
                                linhas_invalidas['Descrição do erro'] = f'VALR_UNIDADE_COMERCIAL menor que {valor_minimo}'
                                df_de_erros = pd.concat([df_de_erros, linhas_invalidas])

                                del linhas_invalidas

                            df_tabela_inteira['DESC_PRODUTO'] = df_tabela_inteira['DESC_PRODUTO'].apply(
                                lambda x: str(x).lower().strip().replace('  ', ' ').replace('\n', '').replace('?',''))

                            df_tabela_inteira = df_tabela_inteira.drop(columns=['VALR_UNIDADE_COMERCIAL'])
                            self.progressBar_2_bebidas.setValue(30)

                            # 3° processo: Remover NF(e) com valor 0
                            self.label_5_bebidas.setText('Removendo linhas.')

                            linhas_unicas = df_tabela_inteira.groupby('CODG_EAN').filter(lambda x: len(x) == 1)
                            linhas_unicas['Descrição do erro'] = 'CODG_EAN unico'

                            df_de_erros = pd.concat([df_de_erros, linhas_unicas])

                            df_tabela_inteira = df_tabela_inteira.groupby('CODG_EAN').filter(lambda x: len(x) > 1)
                            df_tabela_inteira['QTDE_COMERCIAL'] = df_tabela_inteira['QTDE_COMERCIAL'].fillna(0).astype(
                                int)

                            df_tabela_inteira['CODG_EAN'] = df_tabela_inteira.apply(
                                lambda linha_f: corrigeir_gtin(linha_f), axis=1)

                            quant_igual_0 = df_tabela_inteira[df_tabela_inteira['QTDE_COMERCIAL'] == 0].copy()
                            quant_igual_0['Descrição do erro'] = 'QTDE_COMERCIAL igual a 0'

                            df_de_erros = pd.concat([df_de_erros, quant_igual_0])
                            df_tabela_inteira = df_tabela_inteira[df_tabela_inteira['QTDE_COMERCIAL'] != 0]

                            del linhas_unicas, quant_igual_0
                            self.progressBar_2_bebidas.setValue(40)

                            # 4° processo: Criar a coluna do COD_CNAE
                            self.label_5_bebidas.setText('Criando colunas')

                            df_tabela_inteira = pd.merge(df_tabela_inteira, df_fornecedores,
                                                         on='NUMR_INSC_ESTADUAL_EMISSOR', how='left')

                            df_tabela_inteira['DESC_SUBSETOR_SEFAZ'].fillna('DEMAIS_SETORES_MT', inplace=True)
                            df_tabela_inteira['INDICES_DOS_SETORES'].fillna(0, inplace=True)

                            df_tabela_inteira.rename(columns={'DESC_SUBSETOR_SEFAZ': 'DESC_EMISSOR',
                                                     'INDICES_DOS_SETORES': 'ÍNDICES_FORNECEDOR'}, inplace=True)
                            print(f'A pós fornecedores: {len(df_tabela_inteira)}')
                            self.progressBar_2_bebidas.setValue(50)
                            del df_fornecedores

                            # 5° processo: Remover linhas com a quantidade em decimal ou 0 - ok
                            self.label_5_bebidas.setText('Removendo linhas.')
                            # Filtrar linhas com valores iguais a zero ou valores decimais diferentes de zero
                            linhas_invalidas = df_tabela_inteira[
                                df_tabela_inteira['QTDE_COMERCIAL'].astype(str).str.contains(
                                    r'\.0+$|(?<!^)0+\.\d+|^0+$')].copy()

                            linhas_invalidas['Descrição do erro'] = 'QTDE_COMERCIAL não atende à regra'

                            df_de_erros = pd.concat([df_de_erros, linhas_invalidas])
                            df_tabela_inteira = df_tabela_inteira[~df_tabela_inteira.index.isin(linhas_invalidas.index)]
                            print(f'A pó remover valores menoress que 0 {len(df_tabela_inteira)}')
                            del linhas_invalidas
                            self.progressBar_2_bebidas.setValue(60)

                            # 6° processo: Remover CODG_EAN que não estão no arquivo gtin
                            self.label_5_bebidas.setText('Removendo COD_EAN')
                            df_tabela_inteira['CODG_EAN'] = df_tabela_inteira['CODG_EAN'].astype(str)
                            df_gtin['CODG_EAN'] = df_gtin['CODG_EAN'].astype(str)
                            df_filtro['CODG_EAN'] = df_filtro['CODG_EAN'].astype(str)
                            df_tabela_inteira = pd.merge(df_tabela_inteira, df_gtin,
                                                         on='CODG_EAN', how='left')
                            df_tabela_inteira['DESC_PRODUTO_SIMPLIFICADA'].fillna('nan', inplace=True)

                            df_linhas_removidas = df_tabela_inteira[
                                df_tabela_inteira['DESC_PRODUTO_SIMPLIFICADA'].str.lower() == 'nan'].copy()
                            df_tabela_inteira = df_tabela_inteira[
                                df_tabela_inteira['DESC_PRODUTO_SIMPLIFICADA'].str.lower() != 'nan']

                            df_linhas_removidas['Descrição do erro'] = 'gtin não encontrado'
                            df_de_erros = pd.concat([df_de_erros, df_linhas_removidas])

                            del df_linhas_removidas
                            self.progressBar_2_bebidas.setValue(70)
                            print(f'A pós o gtin: {len(df_tabela_inteira)}')

                            # 8° processo: Remover CODG_EAN que não estão no arquivo gtin.json
                            self.label_5_bebidas.setText('Aplicando filtro.')

                            df_excluidos = df_tabela_inteira[
                                df_tabela_inteira['DESC_PRODUTO'].isin(df_filtro['DESC_PRODUTO_IRRELEVANTES'])]

                            # Remover as linhas com LIXO do DataFrame restantes
                            df_tabela_inteira = df_tabela_inteira[~df_tabela_inteira['DESC_PRODUTO'].isin(
                                df_filtro['DESC_PRODUTO_IRRELEVANTES'])]

                            df_excluidos['Descrição do erro'] = 'Reprovado pela descrição reprovar para todos'
                            df_de_erros = pd.concat([df_de_erros, df_excluidos])
                            del df_excluidos
                            self.progressBar_2_bebidas.setValue(80)
                            print(f'Apos reprovar para todos: {len(df_tabela_inteira)}')

                            # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                            # APLICAR FILTRO NUMR_INSC_ESTADUAL_EMISSOR!!
                            # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                            if filtro_iee:
                                if path.exists(".\\confg\\Filtro_IEE.xlsx"):
                                    df_filtro_iee = pd.read_excel(io=".\\confg\\Filtro_IEE.xlsx")
                                else:

                                    df_filtro_iee = pd.DataFrame(columns=['CODG_EAN'])
                                df_merged = df_tabela_inteira.merge(df_filtro_iee[['CODG_EAN']], on='CODG_EAN',
                                                                    how='left', indicator=True)

                                df_linhas_separadas = df_merged[df_merged['_merge'] == 'both'].drop(columns=['_merge'])

                                df_tabela_inteira = df_merged[df_merged['_merge'] == 'left_only'].drop(
                                    columns=['_merge'])

                            # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                            # APLICAR FILTRO RECUSA!!
                            # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
                            if usar_filtro:
                                df_tabela_inteira = df_tabela_inteira.merge(df_filtro[['CODG_EAN', 'DESC_PRODUTO']],
                                                                            on=['CODG_EAN', 'DESC_PRODUTO'], how='left',
                                                                            indicator=True)

                                df_excluidos = df_tabela_inteira[df_tabela_inteira['_merge'] == 'both'].drop(
                                    columns=['_merge'])
                                df_tabela_inteira = df_tabela_inteira[df_tabela_inteira['_merge'] == 'left_only']

                                # Adicionar linhas ao DataFrame df_erros
                                df_excluidos['Descrição do erro'] = 'Reprovado pela descrição reprovar para gtin'
                                df_de_erros = pd.concat([df_de_erros, df_excluidos])
                                del df_excluidos
                                print(f'apos usar filtro: {len(df_tabela_inteira)}')

                            if filtro_iee:
                                df_tabela_inteira = pd.concat([df_tabela_inteira, df_linhas_separadas])
                                del df_linhas_separadas

                            if len(df_tabela_inteira) != 0:
                                print(f'df_tabela_inteira != {len(df_tabela_inteira)}')
                                # 10 processo: Calculando valor unitario
                                self.label_5_bebidas.setText('Calculando valor unitário.')
                                df_gtin = df_gtin.loc[:, ~df_gtin.columns.duplicated()]
                                cont_gtin = df_gtin.set_index('CODG_EAN').T.to_dict('list')
                                df_tabela_inteira[['VALR_UNIDADE_CALC', 'QTDE_COMERCIAL_CALC',
                                                   'QUANT_PONDERAVEIS']] = df_tabela_inteira.apply(
                                    lambda linha: calcular_quantidade_valor_unitario(linha, cont_gtin, valor_minimo,
                                                                                     fitro_valor_minimo), axis=1)

                                # 10° processo: Ordenar grupos
                                print(f'# 10° processo: {len(df_tabela_inteira)}')
                                self.label_5_bebidas.setText('Agrupando valores.')

                            df_de_erros = df_de_erros[['Descrição do erro', 'NUMR_INSC_ESTADUAL_EMISSOR',
                                                       'NOME_FANTASIA_EMISSOR', 'DESC_PRODUTO', 'CODG_EAN',
                                                       'VALR_UNIDADE_COMERCIAL', 'QTDE_COMERCIAL', 'VALR_PRODUTO']]
                            df_tabela_inteira = df_tabela_inteira[
                                        ['ÍNDICES_FORNECEDOR', 'NUMR_INSC_ESTADUAL_EMISSOR',
                                         'NOME_FANTASIA_EMISSOR', 'DESC_EMISSOR',
                                         'DESC_PRODUTO',
                                         'CODG_EAN', 'DESC_PRODUTO_SIMPLIFICADA',
                                         'UNID_COMERCIAL', 'QTDE_COMERCIAL',
                                         'QTDE_COMERCIAL_CALC', 'VALR_UNIDADE_CALC',
                                         'VALR_PRODUTO', 'QUANT_PONDERAVEIS']]

                            self.progressBar_2_bebidas.setValue(90)

                            # 11° processo: Salva o primeiro DataFrame
                            self.label_5_bebidas.setText('Salvando arquivos.')

                            writer = salvar_dataframe_para_excel(arquivo_excel, 'Dados padronizados', df_tabela_inteira)
                            del df_tabela_inteira
                            writer = salvar_dataframe_para_excel(arquivo_excel, 'Tabela de erros', df_de_erros,
                                                                 w=writer)
                            del df_de_erros

                            try:
                                writer.save()
                            except AttributeError:
                                writer._save()
                            del writer

                            self.label_5_bebidas.setText('Tudo Pronto!')
                            self.label_4_bebidas.setText(f' | {hora()} |')
                            self.progressBar_2_bebidas.setValue(100)


                        else:
                            self.msgBox.setText(f"Arquivos {str(caminho_arquivo.split('.')[-1])}ainda não esta pronta!")
                            self.msgBox.exec()
                    else:
                        self.label_5_bebidas.setText('Carregando')
                        self.label_3_bebidas.setText(f'| {hora()}')
                        self.label_4_bebidas.setText('')
                        self.progressBar_2_bebidas.setValue(0)
                        df = pd.read_excel(caminho_arquivo, sheet_name='Dados padronizados', engine='openpyxl')
                        df_de_erros = pd.read_excel(caminho_arquivo, sheet_name='Tabela de erros',
                                                    engine='openpyxl')

                        if path.exists(".\\confg\\Gtin.xlsx"):
                            df_gtin = pd.read_excel(io=".\\confg\\Gtin.xlsx")
                        else:
                            df_gtin = pd.DataFrame(columns=['QTDE', 'VALOR (R$) (PMPF)', 'UNIDADE DE MEDIDA',
                                                            'DESCRIÇÃO', 'GTIN_EAN', '1', '2', '3', '4', '5'])
                        df_gtin = df_gtin.set_index('CODG_EAN').T.to_dict('list')

                        df_de_desvio_padrao = pd.DataFrame(columns=[
                            'Descrição do erro', 'LIMITE_INFERIOR', 'LIMITE_SUPERIOR', 'NOME_FANTASIA_EMISSOR',
                            'NUMR_INSC_ESTADUAL_EMISSOR',
                            'DESC_EMISSOR', 'DESC_PRODUTO',
                            'CODG_EAN', 'DESC_PRODUTO_SIMPLIFICADA', 'QTDE_COMERCIAL_CALC',
                            'VALR_UND_CALC', 'VALR_PRODUTO'
                        ])
                        df_de_fornecedores = pd.DataFrame(columns=['', '', '', '', '', '', '', '', '', '', '', '',
                                                                   'Distribuição e valor médio das vendas por setor',
                                                                   '', '', '', '', '', '', '', ''])
                        tabela_fornecedores = [
                            ['', '', '', '', '', '', '', '', '', 'COMERCIO ATACADISTA DE ALIMENTOS E AFINS',
                             '', 'COMERCIO VAREJISTA DE COMBUSTIVEIS E LUBRIFICANTES', '',
                             'SUPERMERCADOS E ALIMENTOS',
                             '', 'SERVICOS DE ALOJAMENTO E ALIMENTACAO', '', 'DEMAIS SETORES', '', 'Total', ''],
                            ['Ordem', 'GTIN/EAN', 'Descrição', 'Unidade de medida', 'PMPF App', 'pmpf_pauta',
                             'Variação', 'SOMA_QTDE_COMERCIAL', 'Emissores (IE)', 'PMPF(SETOR)', 'Qtde vendida',
                             'PMPF(SETOR)', 'Qtde vendida', 'PMPF(SETOR)', 'Qtde vendida', 'PMPF(SETOR)',
                             'Qtde vendida', 'PMPF(SETOR)', 'Qtde vendida', 'Qtde vendida', 'Total']]
                        df_tabela_resultado = pd.DataFrame(columns=['CONDG_EAN_CALC', 'DESC_SIMP',
                                                                    'CONT_NFE', 'DESABILITADAS (%)',
                                                                    'SOMA_QTDE_COMERCIAL',
                                                                    'VALR_MIN_UND_COMERCIAL',
                                                                    'VALR_MEDIA_UND_COMERCIAL',
                                                                    'VALR_MAX_UND_COMERCIAL', 'AMPLITUDE_UND_CALC',
                                                                    'VAR_UND_CALC',
                                                                    'DESV.P_UND_CALC', 'PMPF App',
                                                                    'SOMA_VALR_PRODUTO',
                                                                    'valor_total_item_percentual'])
                        dicionario_de_pmpf_por_gtin = {}
                        dicionario_de_intervalo = {}
                        tabela_resultado = []
                        df_valores_corretos = pd.DataFrame(columns=df.columns)
                        df_valores_removidos = pd.DataFrame(columns=df_de_desvio_padrao.columns)

                        # Calculando a soma para cada item da coluna CODG_EAN
                        dicionario_por_grupo = df.groupby('CODG_EAN')['VALR_UNIDADE_CALC'].apply(list).to_dict()

                        for a, b in dicionario_por_grupo.items():
                            self.label_5_bebidas.setText(f'Calculando {a}')

                            limites = calcular_intervalo_aceitacao(b)
                            limite_inferior = limites[0]
                            limite_superior = limites[1]
                            dicionario_de_intervalo[a] = [limite_inferior, limite_superior]
                        self.progressBar_2_bebidas.setValue(50)

                        for grupo_nome, grupo_df in df.groupby('CODG_EAN'):
                            limite_inferior, limite_superior = dicionario_de_intervalo[grupo_nome]

                            filtro = grupo_df['VALR_UNIDADE_CALC'].between(limite_inferior, limite_superior)

                            grupo_df_filtrado = grupo_df[filtro].dropna(axis=1, how='all')

                            df_valores_corretos = pd.concat([df_valores_corretos, grupo_df_filtrado],
                                                            ignore_index=True)
                            df_valores_removidos = pd.concat([df_valores_removidos, grupo_df[~filtro]],
                                                             ignore_index=True)

                        self.progressBar_2_bebidas.setValue(60)

                        df = df_valores_corretos.reset_index(drop=True)
                        df_de_desvio_padrao = df_valores_removidos.reset_index(drop=True)
                        filtro = ['NOME_FANTASIA_EMISSOR', 'NUMR_INSC_ESTADUAL_EMISSOR',
                                  'DESC_EMISSOR	DESC_PRODUTO', 'CODG_EAN',
                                  'DESC_PRODUTO_SIMPLIFICADA', 'QTDE_COMERCIAL_CALC', 'VALR_UND_CALC',
                                  'VALR_PRODUTO', 'Índices_Fornecedor', 'COD_CNAE', 'CODG_EAN_TRIB',
                                  'UNID_COMERCIAL', 'QTDE_COMERCIAL', 'QUANT_PONDERAVEIS']
                        df_de_desvio_padrao = df_de_desvio_padrao.filter(items=filtro)
                        df = df.dropna(subset=['DESC_PRODUTO'])
                        del filtro

                        variancia_por_gtin = pd.pivot_table(df, values='VALR_UNIDADE_CALC', index='CODG_EAN',
                                                            aggfunc='var')
                        variancia_desv_p = pd.pivot_table(df, values='VALR_UNIDADE_CALC', index='CODG_EAN',
                                                          aggfunc='std')
                        valor_total_nota = float(df['VALR_PRODUTO'].sum())
                        df = df.groupby('CODG_EAN').filter(lambda x: len(x) > 1)
                        for gtin in df['CODG_EAN'].unique():
                            # Calculando quantidade de linhas para cada aba
                            total_linhas_certas = int((df['CODG_EAN'] == gtin).sum()) + int(
                                (df_de_erros['CODG_EAN'] == gtin).sum()) + int(
                                (df_de_desvio_padrao['CODG_EAN'] == gtin).sum())
                            total_linhas_erradas = int((df_de_erros['CODG_EAN'] == gtin).sum()) + int(
                                (df_de_desvio_padrao['CODG_EAN'] == gtin).sum())
                            percentual = float((total_linhas_erradas * 100) / total_linhas_certas)

                            linha = []
                            valr_produto = float(df.loc[df['CODG_EAN'] == gtin, 'VALR_PRODUTO'].sum())
                            variancia_por_gtin_desv_p = variancia_desv_p.loc[gtin, 'VALR_UNIDADE_CALC']
                            variancia_do_gtin = variancia_por_gtin.loc[gtin, 'VALR_UNIDADE_CALC']
                            cont_nfe = float((df['CODG_EAN'] == gtin).sum())
                            soma_qtde_comercial = float(df.loc[df['CODG_EAN'] == gtin, 'QTDE_COMERCIAL_CALC'].sum())
                            minimo = float(df[df['CODG_EAN'] == gtin]['VALR_UNIDADE_CALC'].min())
                            maximo = float(df[df['CODG_EAN'] == gtin]['VALR_UNIDADE_CALC'].max())
                            media = float(df[df['CODG_EAN'] == gtin]['VALR_UNIDADE_CALC'].mean())
                            valor_total_nota_gtin = float(df.loc[df['CODG_EAN'] == gtin, 'VALR_PRODUTO'].sum())
                            valor_total_item_percentual = float((valor_total_nota_gtin * 100) / valor_total_nota)

                            if (valor_total_nota_gtin != 0) and (soma_qtde_comercial != 0):
                                pmpf = float(valor_total_nota_gtin / soma_qtde_comercial)
                            else:
                                pmpf = 0
                            gtin=str(gtin)
                            linha.append(str(gtin))  # 'CONDG_EAN_CALC'
                            linha.append(str("df_gtin[str(gtin)][0]"))  # 'DESC_SIMP'
                            linha.append(int(cont_nfe))  # 'CONT_NFE'
                            linha.append(float(f'{percentual:.2f}'))
                            linha.append(float(f'{soma_qtde_comercial:.2f}'))  # 'SOMA_QTDE_COMERCIAL
                            linha.append(float(f'{minimo:.2f}'))  # 'VALR_MIN_UND_COMERCIAL'
                            linha.append(float(f'{media:.2f}'))  # 'VALR_MEDIA_UND_COMERCIAL'
                            linha.append(float(f'{maximo:.2f}'))  # VALR_MAX_UND_COMERCIAL'
                            linha.append(float(f'{maximo - minimo:.2f}'))  # 'AMPLITUDE_UND_CALC'
                            linha.append(float(f'{variancia_do_gtin:.2f}'))  # 'VAR_UND_CALC'
                            linha.append(float(f'{variancia_por_gtin_desv_p:.2f}'))  # 'DESV.P_UND_CALC'
                            linha.append(float(f'{pmpf:.2f}'))  # 'PMPF'
                            linha.append(float(f'{valr_produto:2f}'))  # VALR_PRODUTO
                            linha.append(str(f'{valor_total_item_percentual:.2f}%'))  # valor_total_item_percentual
                            dicionario_de_pmpf_por_gtin[gtin] = (
                                [f'{float(pmpf):.2f}', f'{soma_qtde_comercial:.2f}'])
                            self.progressBar_2_bebidas.setValue(80)
                            tabela_resultado.append(linha)
                        lista_de_indces = [1, 2, 3, 4, 5]
                        contagem_por_gtin = df.groupby('CODG_EAN')['NUMR_INSC_ESTADUAL_EMISSOR'].nunique()
                        sequencia = 0
                        for gtin in df['CODG_EAN'].unique():
                            lista = []
                            sequencia += 1
                            soma_total_q_vendida = float(str(dicionario_de_pmpf_por_gtin[str(gtin)][1]))
                            lista.append(int(f'{sequencia}'))  # Ordem
                            lista.append(str(f'{gtin}'))  # GTIN/EAN
                            lista.append(str(df_gtin[int(gtin)][0].strip()))  # Descrição
                            lista.append(str(df_gtin[int(gtin)][1].strip()))  # Unidade de medida
                            lista.append(str(dicionario_de_pmpf_por_gtin[str(gtin)][0]))  # PMPF App
                            pmpf_antigo = float(str(df_gtin[int(gtin)][2]))
                            lista.append(str(pmpf_antigo))
                            pmpf_novo = float(str(dicionario_de_pmpf_por_gtin[str(gtin)][0]))
                            lista.append(f'{str(diferca_de_pmpf(antigo=pmpf_antigo, novo=pmpf_novo))}%')
                            t_per = 0
                            lista.append(soma_total_q_vendida)
                            lista.append(int(contagem_por_gtin[gtin]))  # Emissores (IE)
                            for indice in lista_de_indces:

                                # Soma da coluna 'QTDE_COMERCIAL_CALC' para o valor específico de Índices_Fornecedor
                                soma_qtde_comercial = float(df.loc[(df['CODG_EAN'] == gtin) & (
                                        df['ÍNDICES_FORNECEDOR'] == indice), 'QTDE_COMERCIAL_CALC'].sum())
                                try:
                                    percentual_vendido = round(
                                        float((soma_qtde_comercial * 100) / soma_total_q_vendida), 2)
                                except ZeroDivisionError:
                                    percentual_vendido = 0

                                # Soma da coluna 'VALR_PRODUTO' para o valor específico de Índices_Fornecedor
                                valor_total_nota = float(df.loc[(df['CODG_EAN'] == gtin) & (
                                        df['ÍNDICES_FORNECEDOR'] == indice), 'VALR_PRODUTO'].sum())
                                if (valor_total_nota != 0) and (soma_qtde_comercial != 0):
                                    pmpf = round(float(valor_total_nota / soma_qtde_comercial), 2)
                                else:
                                    pmpf = float(0.00)
                                t_per += percentual_vendido
                                lista.append(pmpf)
                                lista.append(f"{str(percentual_vendido)}%")

                            lista.append(str(dicionario_de_pmpf_por_gtin[str(gtin)][0]))
                            lista.append(f'{t_per:.2f}%')
                            tabela_fornecedores.append(lista)
                            self.progressBar_2_bebidas.setValue(90)
                        dataframes = [pd.DataFrame([lista], columns=df_de_fornecedores.columns) for lista in
                                      tabela_fornecedores]
                        if not dataframes:
                            pass
                        else:
                            df_de_fornecedores = pd.concat(dataframes, ignore_index=True)

                        dataframes = [pd.DataFrame([lista], columns=df_tabela_resultado.columns) for lista in
                                      tabela_resultado]
                        if not dataframes:
                            pass
                        else:
                            df_tabela_resultado = pd.concat(dataframes, ignore_index=True)  # df_tabela_resultado

                        # Analise de quantidade de erros por NUMR_INSC_ESTADUAL_EMISSOR
                        df_analise_quantidade_erros = df_de_erros[
                            ['NUMR_INSC_ESTADUAL_EMISSOR', 'NOME_FANTASIA_EMISSOR', 'VALR_PRODUTO']]
                        df_analise_quantidade_erros = df_analise_quantidade_erros.dropna()

                        lista_num_insc = []
                        lista_nome = []
                        lista_quant = []
                        lista_valor = []
                        for nome_grupo, valor_grupo in df_analise_quantidade_erros.groupby(
                                'NUMR_INSC_ESTADUAL_EMISSOR'):
                            quantidade_erros = int(valor_grupo.shape[0])
                            valor_erros = float(valor_grupo['VALR_PRODUTO'].sum())
                            nome_fantasia = str(valor_grupo['NOME_FANTASIA_EMISSOR'].iloc[0])

                            lista_num_insc.append(nome_grupo)
                            lista_nome.append(nome_fantasia)
                            lista_quant.append(quantidade_erros)
                            lista_valor.append(valor_erros)

                        df_analise_quantidade_erros = pd.DataFrame({
                            'NUMR_INSC_ESTADUAL_EMISSOR': lista_num_insc,
                            'NOME_FANTASIA_EMISSOR': lista_nome,
                            'Quantidade de erros': lista_quant,
                            'Valor total dos erros': lista_valor
                        })

                        self.label_5_bebidas.setText('Salvando tabela.')

                        df = ordenar_grupo(df)
                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo,
                                                             nome_aba='Dados padronizados',
                                                             dataframe=df, w=None)
                        del df
                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo,
                                                             nome_aba='Tabela de erros',
                                                             dataframe=df_de_erros, w=writer)
                        del df_de_erros
                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo,
                                                             nome_aba='DADOS TRATADOS DESVPAD',
                                                             dataframe=df_de_desvio_padrao, w=writer)
                        del df_de_desvio_padrao
                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo, nome_aba='Relatorio',
                                                             dataframe=df_tabela_resultado, w=writer)
                        del df_tabela_resultado
                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo,
                                                             nome_aba='Relatorio fornecedores',
                                                             dataframe=df_de_fornecedores, w=writer)

                        writer = salvar_dataframe_para_excel(nome_arquivo=caminho_arquivo,
                                                             nome_aba='Erros dos fornecedores',
                                                             dataframe=df_analise_quantidade_erros, w=writer)
                        try:
                            writer.save()
                        except AttributeError:
                            writer._save()

                        del writer
                        self.label_5_bebidas.setText('Tudo pronto!')
                        self.label_4_bebidas.setText(f' | {hora()} |')
                        self.progressBar_2_bebidas.setValue(100)
                else:
                    self.msgBox.setText('Por favor, selecione um arquivo e informe um valor mínimo inteiro.')
                    self.msgBox.exec()
            except Exception as erro:
                traceback_str = str(traceback.format_exc())
                registrar_erro(nome_arquivo=r".\\confg\\Erros.txt", texto=str(traceback_str))
                print(erro, traceback_str)
                self.label_5_bebidas.setText('')
                self.label_3_bebidas.setText('')
                self.label_4_bebidas.setText('')
                self.msgBox.exec()
                sys.exit()

    def iniciar_medicamentos(self):
        caminho_dos_arquivos = list(str(self.txt_file_path_medicamentos.text()).split(';'))
        valor_minimo = self.txt_file_path_2_medicamentos.text()
        tarefa = self.radioButton_medicamentos.isChecked()
        usar_filtro = self.checkBox_medicamentos.isChecked()
        remover_duplicatas = self.checkBox2_medicamentos.isChecked()
        filtro_iee = self.checkBox3_medicamentos.isChecked()

        for caminho_arquivo in caminho_dos_arquivos:
            try:
                self.label_5_medicamentos.setText('Projeto ainda em desenvolvimento.')
                 # Projeto ainda em desenvolvimento. Aba de medicamentos ainda está sendo construída.
            except:
                pass


if __name__ == "__main__":
    configurar_ambiente()
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
