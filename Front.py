import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QToolButton, QMenu, QAction, 
                             QHBoxLayout, QVBoxLayout, QWidget, QLabel, QMessageBox, 
                             QPushButton, QLineEdit, QFileDialog)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt  # Adicione esta linha para importar Qt
from atlassian import Jira  # Certifique-se de que a biblioteca atlassian-python-api está instalada

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle('Escolha um número')
        self.setGeometry(100, 100, 400, 400)  # Ajusta a altura da janela após remover evidências
        
        self.widget = QWidget(self)
        self.setCentralWidget(self.widget)

        # Layout principal vertical
        self.main_layout = QVBoxLayout(self.widget)
        
        # Layout horizontal para o campo de entrada e o botão de arquivo
        self.file_layout = QHBoxLayout()
        
        # Campo de entrada para o caminho do arquivo
        self.file_input = QLineEdit(self)
        self.file_input.setPlaceholderText("Caminho do arquivo")  # Adiciona o placeholder
        self.file_layout.addWidget(self.file_input)
        
        # Botão para selecionar o arquivo
        self.file_button = QPushButton('Selecionar Arquivo', self)
        self.file_button.clicked.connect(self.open_file_dialog)
        self.file_layout.addWidget(self.file_button)
        
        # Adicionando o layout do arquivo ao layout principal
        self.main_layout.addLayout(self.file_layout)

        # Layout para os campos de usuário e senha
        self.user_layout = QVBoxLayout()

        # Campo de entrada para o usuário
        self.user_label = QLabel('Usuário:', self)
        self.user_layout.addWidget(self.user_label)
        self.user_input = QLineEdit(self)
        self.user_layout.addWidget(self.user_input)

        # Campo de entrada para a senha
        self.password_layout = QHBoxLayout()
        self.password_label = QLabel('Senha:', self)
        self.user_layout.addWidget(self.password_label)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_layout.addWidget(self.password_input)
        
        # Ícones do botão de "olho"
        self.eye_icon_closed = QIcon('Olho_Fechado.png')
        self.eye_icon_open = QIcon('Olho_Aberto.png')
        
        # Botão de "olho" para ver a senha
        self.eye_button = QToolButton(self)
        self.eye_button.setIcon(self.eye_icon_closed)
        self.eye_button.setCheckable(True)
        self.eye_button.toggled.connect(self.toggle_password)
        self.password_layout.addWidget(self.eye_button)

        self.user_layout.addLayout(self.password_layout)
        
        # Adicionando o layout de usuário e senha ao layout principal
        self.main_layout.addLayout(self.user_layout)

        # Layout horizontal para dividir chamados e login do analista
        self.split_layout = QHBoxLayout()

        # Botão de menu tipo hambúrguer para dividir chamados
        self.split_menuButton = QToolButton(self)
        self.split_menuButton.setText('≡')
        self.split_menuButton.setPopupMode(QToolButton.InstantPopup)

        # Menu associado ao botão de dividir chamados
        self.split_menu = QMenu(self)
        self.split_menuButton.setMenu(self.split_menu)

        # Adicionando ações ao menu de dividir chamados
        split_options = ['Sim', 'Não']
        for option in split_options:
            action = QAction(option, self)
            action.triggered.connect(lambda checked, o=option: self.set_split_option(o))
            self.split_menu.addAction(action)

        self.split_layout.addWidget(QLabel('Dividir chamados:', self))
        self.split_layout.addWidget(self.split_menuButton)

        # Campo de entrada para o login do analista
        self.analyst_input = QLineEdit(self)
        self.analyst_input.setPlaceholderText('Login do analista')
        self.analyst_input.setVisible(False)
        self.split_layout.addWidget(self.analyst_input)

        # Adicionando o layout horizontal de dividir chamados ao layout principal
        self.main_layout.addLayout(self.split_layout)

        # Botão para apresentar a escolha
        self.result_button = QPushButton('Start', self)
        self.result_button.clicked.connect(self.confirm_choice)
        self.main_layout.addWidget(self.result_button)

        # Adiciona um QLabel com a imagem ao final
        self.image_label = QLabel(self)
        self.pixmap = QPixmap('Zé_bonitinho.png')  # Substitua pelo caminho da sua imagem
        self.image_label.setPixmap(self.pixmap)
        self.image_label.setScaledContents(True)  # Ajusta a imagem ao tamanho do QLabel
        self.image_label.setAlignment(Qt.AlignCenter)
        
        self.main_layout.addWidget(self.image_label)

        self.main_layout.addStretch(1)

        self.JIRA_USERNAME = None
        self.senha = None
        self.Querencia_de_dividir = None
        self.Analista02 = None
        self.Arquivo_ler = None  # Variável para armazenar o caminho do arquivo

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.Arquivo_ler = file_path  # Atribui o caminho do arquivo à variável Arquivo_ler
            self.file_input.setText(file_path)
            self.file_input.setToolTip("Caminho do arquivo: " + file_path)  # Adiciona um tooltip informativo
            # Leitura do arquivo XLSX
            self.dataframe = pd.read_excel(file_path)
            print("Arquivo carregado com sucesso!")
            print(self.dataframe)

    def set_split_option(self, option):
        self.Querencia_de_dividir = option
        if option == 'Sim':
            self.analyst_input.setVisible(True)
        else:
            self.analyst_input.setVisible(False)

    def confirm_choice(self):
        self.JIRA_USERNAME = self.user_input.text().strip()
        self.senha = self.password_input.text().strip()
        self.Analista02 = self.analyst_input.text().strip() if self.Querencia_de_dividir == 'Sim' else None

        # Verificar se o caminho do arquivo foi fornecido
        if not self.Arquivo_ler:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, selecione um arquivo.')
            return

        # Verificar se o campo de usuário está preenchido
        if not self.JIRA_USERNAME:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, preencha o campo de usuário.')
            return

        # Verificar se o campo de senha está preenchido
        if not self.senha:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, preencha o campo de senha.')
            return

        # Verificar se o login do analista está preenchido quando necessário
        if self.Querencia_de_dividir == 'Sim' and not self.Analista02:
            QMessageBox.warning(self, 'Login do analista', 'Por favor, preencha o login do analista.')
            return

        # Verifica a conexão com o JIRA
        JIRA_URL = 'https://jira.itamaraty.gov.br'  # Substitua pelo URL do seu JIRA
        try:
            jira = Jira(url=JIRA_URL, username=self.JIRA_USERNAME, password=self.senha)
            projects = jira.get_all_projects()
            if not projects:
                raise Exception("Falha ao obter projetos")
            # Se conseguir se conectar, mostra a caixa de diálogo de confirmação
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle('Confirmação')
            msg_box.setText(f'Dividir chamados: {self.Querencia_de_dividir}\nLogin do analista: {self.Analista02 if self.Querencia_de_dividir == "Sim" else "N/A"}\nCaminho do  arquivo: {self.Arquivo_ler}\nDeseja confirmar?')
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.No)
        
            result = msg_box.exec_()
        
            if result == QMessageBox.Yes:
                self.execute_choice()
        except Exception as e:
            QMessageBox.critical(self, 'Erro de Conexão', f'Não foi possível conectar ao JIRA. Verifique as credenciais e tente novamente.\n\nErro: {e}')
            print(f'Erro ao conectar ao JIRA: {e}')

    def execute_choice(self):
        print(f'Usuário: {self.JIRA_USERNAME}')
        print(f'Dividir chamados: {self.Querencia_de_dividir}')
        print(f'Login do 2º analista: {self.Analista02 if self.Querencia_de_dividir == "Sim" else "N/A"}')
        print(f'Caminho do arquivo: {self.Arquivo_ler}')  # Mostra o caminho do arquivo selecionado
        
        # Adicione a lógica para a execução da escolha aqui.

    def toggle_password(self, checked):
        if checked:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.eye_button.setIcon(self.eye_icon_open)
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.eye_button.setIcon(self.eye_icon_closed)

def main():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
