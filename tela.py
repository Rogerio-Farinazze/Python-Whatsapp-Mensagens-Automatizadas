import locale
import sys
import os
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QFileDialog,
    QMessageBox,
    QLabel,
    QHBoxLayout,
)
# Define o locale para o padrão brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class JanelaPrincipal(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

        # Atributos para armazenar dados
        self.dados_planilha = []
        self.caminho_planilha = ""

    def init_ui(self):
        self.setWindowTitle("Envio de Mensagens pelo WhatsApp")
        self.setGeometry(100, 100, 800, 600)

        # Layout principal
        layout = QVBoxLayout()

        # Botão "Baixar Planilha"
        botao_download = QPushButton("Baixar Planilha Padrão")
        botao_download.clicked.connect(self.baixar_planilha)
        layout.addWidget(botao_download)

        # Botão para upload da planilha
        self.botao_upload = QPushButton("Fazer Upload da Planilha")
        self.botao_upload.clicked.connect(self.carregar_planilha)
        layout.addWidget(self.botao_upload)

        # Tabela para exibir os dados da planilha
        self.tabela = QTableWidget()
        self.tabela.setColumnCount(4)  # Número de colunas: Nome, Telefone, Data Limite, Valor
        self.tabela.setHorizontalHeaderLabels(["Nome", "Telefone", "Data Limite", "Valor"])
        layout.addWidget(self.tabela)

        # Campo para o usuário inserir a mensagem
        layout.addWidget(QLabel("Mensagem personalizada:"))
        self.campo_mensagem = QTextEdit()
        layout.addWidget(self.campo_mensagem)

        # Botão para iniciar o envio
        self.botao_enviar = QPushButton("Iniciar Envio")
        self.botao_enviar.clicked.connect(self.iniciar_envio)
        layout.addWidget(self.botao_enviar)

        # Configura o layout na janela principal
        self.setLayout(layout)

    def baixar_planilha(self):
        # Define o local para salvar o arquivo usando uma janela de seleção
        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar Planilha", "planilha_padrao.xlsx", "Arquivos Excel (*.xlsx)"
        )

        if caminho:
            try:
                # Cria a planilha padrão
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Dados"
                ws.append(["Nome", "Telefone", "Data Limite", "Valor"])  # Cabeçalhos

                # Salva o arquivo no caminho especificado
                wb.save(caminho)

                # Exibe a mensagem de confirmação com a opção de abrir o arquivo
                self.mostrar_mensagem_download_concluido(caminho)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao salvar a planilha: {e}")

    def mostrar_mensagem_download_concluido(self, caminho):
        # Cria a mensagem com botões personalizados
        mensagem = QMessageBox(self)
        mensagem.setIcon(QMessageBox.Information)
        mensagem.setWindowTitle("Download Concluído")
        mensagem.setText(f"Planilha padrão salva em:\n{caminho}")

        # Cria os botões manualmente para personalizar os textos
        botao_fechar = mensagem.addButton("Fechar", QMessageBox.RejectRole)
        botao_abrir = mensagem.addButton("Abrir Arquivo", QMessageBox.AcceptRole)

        # Mostra a mensagem
        mensagem.exec_()

        # Verifica qual botão foi clicado
        if mensagem.clickedButton() == botao_abrir:
            # Abre o arquivo no aplicativo padrão do sistema operacional
            os.startfile(caminho)  # Apenas para Windows; no Linux/Mac, usar subprocess.

    def carregar_planilha(self):
        # Seleciona o arquivo de planilha
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Arquivos Excel (*.xlsx)")
        if not caminho:
            return

        self.caminho_planilha = caminho
        try:
            # Carrega os dados da planilha
            workbook = openpyxl.load_workbook(caminho)
            pagina_clientes = workbook.active

            # Limpa dados anteriores
            self.dados_planilha.clear()
            self.tabela.setRowCount(0)

            # Lê as linhas da planilha e armazena em `dados_planilha`
            for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
                self.dados_planilha.append(linha)

            # Exibe os dados na tabela
            self.tabela.setRowCount(len(self.dados_planilha))
            for i, linha in enumerate(self.dados_planilha):
                for j, valor in enumerate(linha):
                    # Formatar a coluna "Valor" com R$ e duas casas decimais (padrão brasileiro)
                    if j == 3:  # Coluna Valor
                        valor = f"R$ {locale.format_string('%.2f', valor, grouping=True)}" if valor is not None else "R$ 0,00"
                    self.tabela.setItem(i, j, QTableWidgetItem(str(valor)))

            QMessageBox.information(self, "Sucesso", "Planilha carregada com sucesso!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar a planilha: {e}")

    def iniciar_envio(self):
        # Verifica se a planilha foi carregada
        if not self.dados_planilha:
            QMessageBox.warning(self, "Aviso", "Carregue uma planilha antes de iniciar o envio.")
            return

        # Obtém a mensagem personalizada
        mensagem_base = self.campo_mensagem.toPlainText()
        if not mensagem_base:
            QMessageBox.warning(self, "Aviso", "Insira a mensagem personalizada antes de iniciar o envio.")
            return

        # Adiciona o rodapé à mensagem
        rodape = "\n\nSiga nas redes sociais @Conectadomogi\nAcesse o site https://colegioconectado.com.br"
        mensagem_base += rodape

        # Inicia o processo de envio
        for linha in self.dados_planilha:
            nome, telefone, data_limite, valor_devido = linha
            mensagem = mensagem_base.format(
                nome=nome, telefone=telefone, data_limite=data_limite, valor_devido=f"R$ {locale.format_string('%.2f', valor_devido, grouping=True)}" if valor_devido is not None else "R$ 0,00"
            )

            try:
                # Envia a mensagem pelo WhatsApp
                link_whats = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
                webbrowser.open(link_whats)
                sleep(6)
                pyautogui.press('esc')
                sleep(2)
                pyautogui.press('enter')
                sleep(2)
                pyautogui.hotkey('ctrl', 'w')
                sleep(2)
            except Exception as e:
                print(f"Erro ao enviar mensagem para {nome}: {e}")
                with open('erros.csv', 'a+', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'\r\n{nome},{telefone}')

        QMessageBox.information(self, "Envio Concluído", "Mensagens enviadas com sucesso!")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Instancia a janela principal
    janela = JanelaPrincipal()
    janela.show()

    # Executa o loop da aplicação
    sys.exit(app.exec_())
