import sys
import gspread
import unicodedata
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QDateEdit, QMessageBox
)
from PyQt6.QtCore import Qt, QDate
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date, timedelta
from fpdf import FPDF
import matplotlib.pyplot as plt
import os


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = r"C:\\CLOUD\\PROJETOS\\PLANILHA AUTOMATIZADA\\Arquivo JSON\\calm-scarab-455514-n9-25ed6bba9bde.json"
SHEET_ID = "1k_sFMOceBQodz6YyLO4299ishthgy1jiwopz9MgrIO4"


def normalizar(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').lower().strip()


def calcular_dias_corridos(data_str):
    try:
        data_chamado = datetime.strptime(data_str, "%d/%m/%Y").date()
        return (date.today() - data_chamado).days
    except Exception as e:
        print(f"Erro ao calcular dias corridos: {e}")
        return "-"


def calcular_dias_uteis(data_str):
    """
    Calcula o número de dias úteis (segunda a sexta) entre a data fornecida e a data atual.
    """
    try:
        data_chamado = datetime.strptime(data_str, "%d/%m/%Y").date()
        delta = date.today() - data_chamado
        dias_uteis = 0

        for i in range(delta.days + 1):
            dia_atual = data_chamado + timedelta(days=i)
            if dia_atual.weekday() < 5:  # Segunda a sexta-feira (0 a 4)
                dias_uteis += 1

        return dias_uteis
    except Exception as e:
        print(f"Erro ao calcular dias úteis: {e}")
        return "-"


class PDFComPaginacao(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")


class EquipamentoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Controle de Equipamentos")
        self.setGeometry(100, 100, 1200, 700)

        # Layout principal
        layout_principal = QVBoxLayout()

        # Seção de Entrada de Dados
        entrada_layout = QVBoxLayout()
        labels = ["Tipo Equipamento", "Tombamento", "Origem", "Endereco", "Numero O.S", "TICKET TCORP", "Situacao", "Data Chamado", "Data Finalização Chamado"]
        self.inputs = {}
        for label in labels:
            hbox = QHBoxLayout()
            lbl = QLabel(label, self)
            lbl.setStyleSheet("font-weight: bold;")
            hbox.addWidget(lbl)
            input_field = QLineEdit(self)
            if label in ["Data Chamado", "Data Finalização Chamado"]:
                input_field.setInputMask("00/00/0000")
            hbox.addWidget(input_field)
            entrada_layout.addLayout(hbox)
            self.inputs[label] = input_field

        # Botão Enviar
        self.btn_enviar = QPushButton("Enviar para Planilha", self)
        self.btn_enviar.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_enviar.clicked.connect(self.enviar_dados)
        entrada_layout.addWidget(self.btn_enviar)

        layout_principal.addLayout(entrada_layout)

        # Seção de Filtros e Relatórios
        filtro_layout = QHBoxLayout()
        self.date_inicio = QDateEdit()
        self.date_inicio.setDisplayFormat("dd/MM/yyyy")
        self.date_inicio.setDate(QDate.currentDate().addMonths(-1))
        filtro_layout.addWidget(QLabel("Início:"))
        filtro_layout.addWidget(self.date_inicio)

        self.date_fim = QDateEdit()
        self.date_fim.setDisplayFormat("dd/MM/yyyy")
        self.date_fim.setDate(QDate.currentDate())
        filtro_layout.addWidget(QLabel("Fim:"))
        filtro_layout.addWidget(self.date_fim)

        self.btn_pdf = QPushButton("Gerar Relatório PDF")
        self.btn_pdf.setStyleSheet("background-color: #008CBA; color: white; font-weight: bold;")
        self.btn_pdf.clicked.connect(self.gerar_pdf)
        filtro_layout.addWidget(self.btn_pdf)

        layout_principal.addLayout(filtro_layout)

        # Label de Chamados Críticos
        self.label_chamados_criticos = QLabel("Chamados Críticos: 0", self)
        self.label_chamados_criticos.setStyleSheet("color: red; font-weight: bold;")
        layout_principal.addWidget(self.label_chamados_criticos)

        # Botões de Ordenação
        ordenacao_layout = QHBoxLayout()
        btn_ordenar_az = QPushButton("Ordenar A-Z (Tipo)")
        btn_ordenar_az.setStyleSheet("background-color: #f44336; color: white;")
        btn_ordenar_az.clicked.connect(lambda: self.ordenar_tabela(coluna=0, ordem="asc"))
        ordenacao_layout.addWidget(btn_ordenar_az)

        btn_ordenar_za = QPushButton("Ordenar Z-A (Tipo)")
        btn_ordenar_za.setStyleSheet("background-color: #f44336; color: white;")
        btn_ordenar_za.clicked.connect(lambda: self.ordenar_tabela(coluna=0, ordem="desc"))
        ordenacao_layout.addWidget(btn_ordenar_za)

        layout_principal.addLayout(ordenacao_layout)

        # Tabela
        self.tabela = QTableWidget()
        layout_principal.addWidget(self.tabela)

        self.setLayout(layout_principal)
        self.carregar_dados_tabela()

    def enviar_dados(self):
        dados = [self.inputs[label].text() for label in self.inputs]
        if all(dados):
            try:
                creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPES)
                client = gspread.authorize(creds)
                sheet = client.open_by_key(SHEET_ID).sheet1

                # Verifica se a coluna extra já está no cabeçalho
                header = sheet.row_values(1)
                if len(header) < 9:
                    header.append("Data Finalização Chamado")
                    sheet.update('A1:I1', [header])

                sheet.append_row(dados)
                QMessageBox.information(self, "Sucesso", "✅ Dados enviados com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"❌ Erro ao enviar dados: {e}")
            for input_field in self.inputs.values():
                input_field.clear()
            self.carregar_dados_tabela()

    def carregar_dados_tabela(self):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPES)
            client = gspread.authorize(creds)
            sheet = client.open_by_key(SHEET_ID).sheet1
            self.dados = sheet.get_all_values()

            if self.dados and isinstance(self.dados[0][0], str) and self.dados[0][0].lower() in ["tipo", "tipo equipamento"]:
                self.dados = self.dados[1:]

            self.tabela.setRowCount(len(self.dados))
            self.tabela.setColumnCount(9)
            self.tabela.setHorizontalHeaderLabels([
                "Tipo", "Tombamento", "Origem", "Endereço", "Nº O.S", "Ticket", "Situação", "Data Chamado", "Data Finalização Chamado"
            ])

            chamados_criticos = 0

            for row_idx, row_data in enumerate(self.dados):
                situacao = row_data[6].strip().lower()
                data_chamado = row_data[7] if len(row_data) > 7 else ""
                try:
                    dias_uteis = calcular_dias_uteis(data_chamado)
                except:
                    dias_uteis = 0

                for col_idx in range(9):
                    valor = row_data[col_idx] if col_idx < len(row_data) else ""
                    item = QTableWidgetItem(str(valor))

                    # Destacar a linha inteira com base na situação e nos dias úteis
                    if situacao == "concluido":
                        item.setBackground(Qt.GlobalColor.green)  # Concluído
                    elif situacao != "concluido":
                        if dias_uteis > 4:
                            item.setBackground(Qt.GlobalColor.red)  # Crítico
                            chamados_criticos += 1 if col_idx == 0 else 0  # Conta apenas uma vez por linha
                        else:
                            item.setBackground(Qt.GlobalColor.yellow)  # Em andamento

                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.tabela.setItem(row_idx, col_idx, item)

            # Atualiza o label de chamados críticos
            self.label_chamados_criticos.setText(f"Chamados Críticos: {chamados_criticos}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar dados do Google Sheets: {e}")

    def ordenar_tabela(self, coluna, ordem):
        """
        Ordena a tabela pela coluna especificada.
        :param coluna: Índice da coluna a ser ordenada.
        :param ordem: "asc" para ascendente, "desc" para descendente.
        """
        try:
            # Extrai os dados da tabela
            dados_atuais = []
            for row in range(self.tabela.rowCount()):
                linha = []
                for col in range(self.tabela.columnCount()):
                    item = self.tabela.item(row, col)
                    linha.append(item.text() if item else "")
                dados_atuais.append(linha)

            # Ordena os dados
            reverse = True if ordem == "desc" else False
            dados_ordenados = sorted(dados_atuais, key=lambda x: x[coluna].lower(), reverse=reverse)

            # Atualiza a tabela com os dados ordenados
            for row_idx, row_data in enumerate(dados_ordenados):
                for col_idx, valor in enumerate(row_data):
                    item = QTableWidgetItem(str(valor))
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.tabela.setItem(row_idx, col_idx, item)

            # Reaplica as cores após a ordenação
            self.aplicar_cores()

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ordenar tabela: {e}")

    def aplicar_cores(self):
        """
        Aplica cores de destaque às linhas da tabela com base nas regras definidas.
        """
        chamados_criticos = 0

        for row_idx in range(self.tabela.rowCount()):
            situacao_item = self.tabela.item(row_idx, 6)  # Coluna "Situação"
            data_chamado_item = self.tabela.item(row_idx, 7)  # Coluna "Data Chamado"

            if situacao_item and data_chamado_item:
                situacao = situacao_item.text().strip().lower()
                data_chamado = data_chamado_item.text()
                try:
                    dias_uteis = calcular_dias_uteis(data_chamado)
                except:
                    dias_uteis = 0

                for col_idx in range(self.tabela.columnCount()):
                    item = self.tabela.item(row_idx, col_idx)

                    if situacao == "concluido":
                        item.setBackground(Qt.GlobalColor.green)  # Concluído
                    elif situacao != "concluido":
                        if dias_uteis > 4:
                            item.setBackground(Qt.GlobalColor.red)  # Crítico
                            chamados_criticos += 1 if col_idx == 0 else 0  # Conta apenas uma vez por linha
                        else:
                            item.setBackground(Qt.GlobalColor.yellow)  # Em andamento

        # Atualiza o label de chamados críticos
        self.label_chamados_criticos.setText(f"Chamados Críticos: {chamados_criticos}")

    def gerar_pdf(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        caminho = os.path.join(desktop, f"relatorio_com_sessoes_{datetime.now().strftime('%d%m%Y_%H%M%S')}.pdf")
        pdf = PDFComPaginacao()
        pdf.add_page()

        # Cabeçalho personalizado
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "Relatório de Equipamentos", ln=True, align="C")
        pdf.set_font("Arial", "I", 10)
        pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True, align="C")
        pdf.ln(10)

        # Resumo geral
        inicio = datetime.strptime(self.date_inicio.text(), "%d/%m/%Y").date()
        fim = datetime.strptime(self.date_fim.text(), "%d/%m/%Y").date()
        dados_filtrados = [row for row in self.dados if inicio <= datetime.strptime(row[7], "%d/%m/%Y").date() <= fim]

        total_chamados = len(dados_filtrados)
        chamados_concluidos = sum(1 for row in dados_filtrados if row[6].strip().lower() == "concluido")
        chamados_criticos = sum(1 for row in dados_filtrados if calcular_dias_uteis(row[7]) > 4 and row[6].strip().lower() != "concluido")
        chamados_em_andamento = total_chamados - chamados_concluidos - chamados_criticos

        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Resumo Geral", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 10, f"Período: {inicio.strftime('%d/%m/%Y')} até {fim.strftime('%d/%m/%Y')}", ln=True)
        pdf.cell(0, 10, f"Total de Chamados: {total_chamados}", ln=True)
        pdf.cell(0, 10, f"Concluídos: {chamados_concluidos}", ln=True)
        pdf.cell(0, 10, f"Críticos: {chamados_criticos}", ln=True)
        pdf.cell(0, 10, f"Em Andamento: {chamados_em_andamento}", ln=True)
        pdf.ln(15)

        # Gráfico de pizza
        fig, ax = plt.subplots(figsize=(6, 3))
        labels = ["Concluídos", "Críticos", "Em Andamento"]
        sizes = [chamados_concluidos, chamados_criticos, chamados_em_andamento]
        colors = ["green", "red", "yellow"]
        ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors)
        ax.axis("equal")
        grafico_path = os.path.join(os.path.expanduser("~"), "Desktop", "grafico_pizza.png")
        plt.savefig(grafico_path)
        plt.close()

        pdf.image(grafico_path, x=10, y=pdf.get_y(), w=180)
        pdf.ln(60)  # Espaço após o gráfico

        # Detalhes por Tipo de Equipamento
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Detalhes por Tipo de Equipamento", ln=True)
        pdf.set_font("Arial", size=10)

        agrupado_por_tipo = {}
        for row in dados_filtrados:
            tipo = row[0].strip().title()
            agrupado_por_tipo.setdefault(tipo, []).append(row)

        for tipo, linhas in agrupado_por_tipo.items():
            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 10, f"Tipo: {tipo}", ln=True)
            pdf.set_font("Arial", size=10)

            headers = ["Tombamento", "Origem", "Situação", "Data Chamado", "Finalização"]
            larguras = [30, 30, 30, 30, 30]

            for i, h in enumerate(headers):
                pdf.cell(larguras[i], 10, h, border=1)
            pdf.ln()

            for row in linhas:
                finalizacao = row[8] if len(row) > 8 and row[8] else ""
                dados_linha = [row[1], row[2], row[6], row[7], finalizacao]
                for i, val in enumerate(dados_linha):
                    pdf.cell(larguras[i], 10, str(val)[:25], border=1)
                pdf.ln()

            pdf.ln(5)

        # Detalhes por Origem
        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Detalhes por Origem", ln=True)
        pdf.set_font("Arial", size=10)

        agrupado_por_origem = {}
        for row in dados_filtrados:
            origem = row[2].strip().title()
            agrupado_por_origem.setdefault(origem, []).append(row)

        for origem, linhas in agrupado_por_origem.items():
            pdf.set_font("Arial", "B", 11)
            pdf.cell(0, 10, f"Origem: {origem}", ln=True)
            pdf.set_font("Arial", size=10)

            headers = ["Tipo", "Tombamento", "Situação", "Data Chamado", "Finalização"]
            larguras = [30, 30, 30, 30, 30]

            for i, h in enumerate(headers):
                pdf.cell(larguras[i], 10, h, border=1)
            pdf.ln()

            for row in linhas:
                finalizacao = row[8] if len(row) > 8 and row[8] else ""
                dados_linha = [row[0], row[1], row[6], row[7], finalizacao]
                for i, val in enumerate(dados_linha):
                    pdf.cell(larguras[i], 10, str(val)[:25], border=1)
                pdf.ln()

            pdf.ln(5)

        pdf.output(caminho)
        QMessageBox.information(self, "Sucesso", f"📄 Relatório PDF salvo: {caminho}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EquipamentoApp()
    window.show()
    sys.exit(app.exec())