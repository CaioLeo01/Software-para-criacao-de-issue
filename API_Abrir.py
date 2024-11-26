import math
import pandas as pd
import requests
import json
from atlassian import Jira
from PyQt5.QtWidgets import (QApplication, QMainWindow, QToolButton, QMenu, QAction, 
                             QHBoxLayout, QVBoxLayout, QWidget, QLabel, QMessageBox, 
                             QPushButton, QLineEdit, QFileDialog)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
import sys
from requests.auth import HTTPBasicAuth
import subprocess
from IPython.display import display

# Configurações do Jira
SERVICE_DESK_ID = "1"
JIRA_URL = "https://jira.itamaraty.gov.br"


novo_valor_enterprise = 'N3'


# Inicializa um DataFrame vazio para armazenar as issue keys
df_tabela_chamados = pd.DataFrame(columns=['Chamado'])

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle('Open tickets')
        self.setGeometry(100, 100, 400, 400)
        
        self.widget = QWidget(self)
        self.setCentralWidget(self.widget)

        self.main_layout = QVBoxLayout(self.widget)
        self.file_layout = QHBoxLayout()
        
        self.file_input = QLineEdit(self)
        self.file_input.setPlaceholderText("File path")
        self.file_layout.addWidget(self.file_input)
        
        self.file_button = QPushButton('Select file', self)
        self.file_button.clicked.connect(self.open_file_dialog)
        self.file_layout.addWidget(self.file_button)
        
        self.main_layout.addLayout(self.file_layout)

        self.user_layout = QVBoxLayout()
        self.user_label = QLabel('Usuário:', self)
        self.user_layout.addWidget(self.user_label)
        self.user_input = QLineEdit(self)
        self.user_layout.addWidget(self.user_input)

        self.password_layout = QHBoxLayout()
        self.password_label = QLabel('Senha:', self)
        self.user_layout.addWidget(self.password_label)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_layout.addWidget(self.password_input)
        
        self.eye_icon_closed = QIcon('Olho_Fechado.png')
        self.eye_icon_open = QIcon('Olho_Aberto.png')
        self.eye_button = QToolButton(self)
        self.eye_button.setIcon(self.eye_icon_closed)
        self.eye_button.setCheckable(True)
        self.eye_button.toggled.connect(self.toggle_password)
        self.password_layout.addWidget(self.eye_button)

        self.user_layout.addLayout(self.password_layout)
        self.main_layout.addLayout(self.user_layout)

        self.labels_layout = QVBoxLayout()
        self.labels_label = QLabel('Labels:', self)
        self.labels_layout.addWidget(self.labels_label)
        self.labels_input = QLineEdit(self)
        self.labels_layout.addWidget(self.labels_input)
        self.main_layout.addLayout(self.labels_layout)

        self.split_layout = QHBoxLayout()
        self.split_menuButton = QToolButton(self)
        self.split_menuButton.setText('≡')
        self.split_menuButton.setPopupMode(QToolButton.InstantPopup)

        self.split_menu = QMenu(self)
        self.split_menuButton.setMenu(self.split_menu)

        split_options = ['Sim', 'Não']
        for option in split_options:
            action = QAction(option, self)
            action.triggered.connect(lambda checked, o=option: self.set_split_option(o))
            self.split_menu.addAction(action)

        self.split_layout.addWidget(QLabel('Dividir chamados:', self))
        self.split_layout.addWidget(self.split_menuButton)

        self.analyst_input = QLineEdit(self)
        self.analyst_input.setPlaceholderText('Login do analista')
        self.analyst_input.setVisible(False)
        self.split_layout.addWidget(self.analyst_input)

        self.main_layout.addLayout(self.split_layout)

        self.result_button = QPushButton('Start', self)
        self.result_button.clicked.connect(self.confirm_choice)
        self.main_layout.addWidget(self.result_button)

        self.close_tickets_button = QPushButton('Close Tickets', self)
        self.close_tickets_button.clicked.connect(self.close_tickets)
        self.main_layout.addWidget(self.close_tickets_button)

        self.image_label = QLabel(self)
        self.pixmap = QPixmap('Zé_bonitinho.png')
        self.image_label.setPixmap(self.pixmap)
        self.image_label.setScaledContents(True)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setVisible(False)
        
        self.main_layout.addWidget(self.image_label)
        self.main_layout.addStretch(1)

        self.developer_label = QLabel('Desenvolvido por Caio Manoel©2024', self)
        self.developer_label.setAlignment(Qt.AlignCenter)
        self.main_layout.addWidget(self.developer_label)

        self.JIRA_USERNAME = None
        self.senha = None
        self.Querencia_de_dividir = None
        self.Analista02 = None
        self.Arquivo_ler = None
        self.nome_arquivo = "chamados_criados.xlsx"  # Defina o nome do arquivo aqui

        self.c = 0
        self.Limitador_analista = 0

    def open_file_dialog(self):
      file_path, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo", "", "Excel Files (*.xlsx);;All Files (*)")
      if file_path:
            self.Arquivo_ler = file_path
            self.file_input.setText(file_path)
            self.file_input.setToolTip("Caminho do arquivo: " + file_path)
            self.dataframe = pd.read_excel(file_path)
            self.tabela = pd.read_excel(self.Arquivo_ler)

            self.Num_linha = self.tabela.shape[0] - 1
            self.Num_linha01 = math.ceil(self.Num_linha / 2)
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
        self.labels = [label.strip() for label in self.labels_input.text().split(',')]

        if not self.Arquivo_ler:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, selecione um arquivo.')
            return

        if not self.JIRA_USERNAME:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, preencha o campo de usuário.')
            return

        if not self.senha:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, preencha o campo de senha.')
            return

        if not self.labels:
            QMessageBox.warning(self, 'Campo obrigatório', 'Por favor, preencha o campo de labels.')
            return

        if self.Querencia_de_dividir == 'Sim' and not self.Analista02:
            QMessageBox.warning(self, 'Login do analista', 'Por favor, preencha o login do analista.')
            return

        try:
            jira = Jira(url='https://jira.itamaraty.gov.br', username=self.JIRA_USERNAME, password=self.senha)
            projects = jira.get_all_projects()
            if not projects:
                raise Exception("Falha ao obter projetos")
            
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle('Confirmação')
            msg_box.setText(f'Dividir chamados: {self.Querencia_de_dividir}\nLogin do analista: {self.Analista02 if self.Querencia_de_dividir == "Sim" else "N/A"}\nCaminho do arquivo: {self.Arquivo_ler}\nLabels: {self.labels}\nDeseja confirmar?')
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.No)
        
            result = msg_box.exec_()
        
            if result == QMessageBox.Yes:
                self.execute_choice()
        except Exception as e:
            QMessageBox.critical(self, 'Erro de Conexão', f'Não foi possível conectar ao JIRA. Verifique as credenciais e tente novamente.\n\nErro: {e}')
            print(f'Erro ao conectar ao JIRA: {e}')

    def toggle_password(self, checked):
        if checked:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.eye_button.setIcon(self.eye_icon_open)
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.eye_button.setIcon(self.eye_icon_closed)

    def close_tickets(self):
    # Exibe uma caixa de diálogo para confirmar se o usuário deseja encerrar os chamados
        reply = QMessageBox.question(
            self, 
            'Confirmação', 
            "Deseja encerrar os chamados?", 
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Se o usuário clicar em "Sim", chama a função para encerrar os chamados
            self.Encerrar_chamado()
        else:
            # Se o usuário clicar em "Não", exibe uma mensagem e não executa a função
            print("Encerramento dos chamados cancelado pelo usuario.")


    def execute_choice(self):
        print("Executando escolha com os seguintes dados:")
        print(f"Usuário: {self.JIRA_USERNAME}")
        print(f"Senha: {'*' * len(self.senha)}")
        print(f"Arquivo: {self.Arquivo_ler}")
        print(f"Dividir chamados: {self.Querencia_de_dividir}")
        print(f"Login do analista: {self.Analista02}")
        print(f"Labels: {self.labels}")

    def execute_choice(self):
        Analista = self.JIRA_USERNAME
        global df_tabela_chamados
        
        self.image_label.setVisible(True)
        jira = Jira(url=JIRA_URL, username=self.JIRA_USERNAME, password=self.senha)

        
        
        while self.c <= self.Num_linha:
            summary = f"{self.tabela.loc[self.c, 'Summary']}"
            description = f"{self.tabela.loc[self.c, 'Description']}"
            comentario = f"{self.tabela.loc[self.c, 'Comentario']}"
            Item = f"{self.tabela.loc[self.c, 'Item de catalogo']}"
            Querencia_linkar  = f"{self.tabela.loc[self.c, 'Deseja linkar']}"
            Mud = f"{self.tabela.loc[self.c, 'Mud']}"
            Querencia_de_evidencia = f"{self.tabela.loc[self.c, 'Deseja evidencia']}"
            evidencia = f"{self.tabela.loc[self.c, 'Evidencia']}"

            def Escolha_Do_item(Item):
                item_dict = {     
                    #Armazenamento________________________________________________________________________________________________________
                    'N3.1.3.99 Outras Atividades - Sustentar': "692",
                    
                    'N3.1.3.12 Registro de outras Demandas Rotineiras': "691",
                    
                    'N3.1.3.11 Registro de problemas para a equipe de Suporte à Armazenamento': "690",
                    
                    'N3.1.3.10 Registro de incidentes para a equipe de Suporte à Armazenamento': "689",
                    
                    'N3.1.3.9 Reestabelecer funcionamento do armazenamento devido à indisponibilidade/lentidão do serviço': "688",
                    
                    'N3.1.3.8 Reestabelecer funcionamento do VMWare ESXI devido à indisponibilidade/lentidão do serviço': "687",
                    
                    'N3.1.3.7 Análise de logs do Vcenter': "686",
                    
                    'N3.1.3.6 Acompanhamento de Release Notes': "685",
                    
                    'N3.1.3.5 Monitoramento e análise de performance, capacidade e consumo dos recursos de hardware do storage': "684",
                    
                    'N3.1.3.4 Análise e verificação de capacidade para LUN, Volumes e Aggregate': "683",
                    
                    'N3.1.3.3 Análise de relatórios do Storage': "682",
                    
                    'N3.1.3.2 Abertura de chamado para suporte / garantia com fornecedor e acompanhamento': "681",
                    
                    'N3.1.3.1 Elaboração de relatório técnico': "680",
                    
                    'N3.1.2.99 Outras Atividades - Configurar': "679",
                    
                    'N3.1.2.12 Update de ambiente virtual / Servidor físico': "678",
                    
                    'N3.1.2.11 Update de ambiente virtual / Virtual Appliance / VMTools': "677",
                    
                    'N3.1.2.10 Reconfiguração de máquina virtual existente': "676",
                    
                    'N3.1.2.9 Recuperação de Máquinas Virtuais': "675",
                    
                    'N3.1.2.8 Criação, configuração ou exclusão de Agregate': "674",
                    
                    'N3.1.2.7 Configuração de Qtree no storage': "673",
                    
                    'N3.1.2.6 Configuração para aumento de tamanho LUN/Volume': "672",
                    
                    'N3.1.2.5 Criação ou configuração de cotas de armazenamento': "671",
                    
                    'N3.1.2.4 Configuração de NFS/CIFS': "670",
                    
                    'N3.1.2.3 Configuração de Initiator Group e Snapdrive': "669",
                    
                    'N3.1.2.2 Configuração de Desduplicação': "668",
                    
                    'N3.1.2.1 Criação, Configuração ou Exclusão de LUN / Volume': "667",
                    
                    'N3.1.1.99 Outras Atividades - Instalar': "666",
                    
                    'N3.1.1.4 Criação de clone de máquina virtual': "665",
                    
                    'N3.1.1.3 Criação / Exclusão de máquina virtual': "664",
                    
                    'N3.1.1.2 Atualização de versão de softwares das soluções de armazenamento': "663",
                    
                    'N3.1.1.1 Atualização de versão de softwares das soluções de virtualização': "662",
                    
                    #Backup/Restore_________________________________________________________________________________________________________

                    'N3.2.3.99 Outras Atividades - Sustentar': "715",
                    
                    'N3.2.3.12 Registro de outras Demandas Rotineiras': "714",
                    
                    'N3.2.3.11 Registro de problemas para a equipe de Suporte à Armazenamento': "713",
                    
                    'N3.2.3.10 Registro de outros incidentes para a equipe de Suporte à Backup/Restauração': "712",
                    
                    'N3.2.3.9 Investigar as vulnerabilidades de backup/restauração de dados e propor ações de mitigação': "711",
                    
                    'N3.2.3.8 Reestabelecer funcionamento do backup devido à indisponibilidade/lentidão do serviço': "710",
                    
                    'N3.2.3.7 Abertura de chamado para suporte / garantia com fornecedor e acompanhamento': "709",
                    
                    'N3.2.3.6 Análise de logs das rotinas de backup e restore - RMI': "708",
                    
                    'N3.2.3.5 Limpeza dos drives das bibliotecas de fitas': "707",
                    
                    'N3.2.3.4 Import / Export de fitas': "706",
                    
                    'N3.2.3.3 Teste de backup e restore': "705",
                    
                    'N3.2.3.2 Análise de logs das rotinas de backup e restore - SERE': "704",
                    
                    'N3.2.3.1 Elaboração relatório técnico': "703",
                    
                    'N3.2.2.99 Outras Atividades - Configurar': "702",
                    
                    'N3.2.2.6 Configuração de ferramenta de backup Snapmanager for Exchange': "701",
                    
                    'N3.2.2.5 Recuperação de itens do correio eletronico - Exchange SMBR': "700",
                    
                    'N3.2.2.4 Recuperação de Bases de Dados - SQL Server': "699",
                    
                    'N3.2.2.3 Recuperação de arquivos - Servidor de Arquivos': "698",
                    
                    'N3.2.2.2 Configuração de políticas para backup / restore': "697",
                    
                    'N3.2.2.1 Configuração de JOB de Backup': "696",
                    
                    'N3.2.1.99 Outras Atividades - Instalar': "695",
                    
                    'N3.2.1.2 Instalação de ferramenta de backup Snapmanager for Exchange': "694",

                    'N3.2.1.1 Atualização de versão do software da solução de backup': "693",

                    #Serv.Linux_______________________________________________________________________________________________________________________

                    'N3.6.3.99 Outras Atividades - Sustentar': "751",

                    'N3.6.3.17 Registro de outras Demandas Rotineiras': "750",

                    'N3.6.3.16 Identificação e análise de problemas (troubleshooting) para a equipe de Suporte a Servidores Linux': "749",

                    'N3.6.3.15 Registro de outros incidentes para a equipe de Suporte a Servidores Linux': "748",

                    'N3.6.3.14 Investigar as vulnerabilidades de SO Linux e propor ações de mitigação': "747",

                    'N.3.6.3.13 Servidor – Alerta de memória': "746",

                    'N.3.6.3.12 Servidor – Alerta de disco': "745",

                    'N3.6.3.11 Servidor – Alerta de CPU': "744",

                    'N3.6.3.10 Reestabelecer funcionamento de Containers devido à indisponibilidade/lentidão do serviço': "743",

                    'N3.6.3.9 Reestabelecer funcionamento de Linux Server devido à indisponibilidade/lentidão do serviço': "742",

                    'N3.6.3.8 Reestabelecer funcionamento de servidor de aplicação devido à indisponibilidade/lentidão do serviço': "741",

                    'N3.6.3.7 Reestabelecer funcionamento da Docker Swarm devido à indisponibilidade/lentidão do serviço': "740",

                    'N3.6.3.6 Reestabelecer funcionamento da ferramenta de monitoramento devido à indisponibilidade/lentidão do serviço': "739",

                    'N3.6.3.5 Reestabelecer funcionamento do Foreman devido à indisponibilidade/lentidão do serviço': "738",

                    'N3.6.3.4 Reestabelecer funcionamento do Puppet devido à indisponibilidade/lentidão do serviço': "737",

                    'N3.6.3.3 Reestabelecer funcionamento do Spacewalk devido à indisponibilidade/lentidão do serviço': "736",

                    'N3.6.3.2 Reestabelecer funcionamento do Syslog/LogAnalizer devido à indisponibilidade/lentidão do serviço': "735",

                    'N3.6.3.1 Elaboração de relatório técnico': "734",

                    'N3.6.2.99 Outras Atividades - Configurar': "733",

                    'N3.6.2.11 Escrita de novo código no gerenciamento de configurações dos sistemas e serviços do parque UNIX/LINUX': "732",

                    'N3.6.2.10 Configuração de Template global já existente na ferramenta de gerenciamento das configurações dos sistemas e serviços do parque UNIX/LINUX': "731",

                    'N3.6.2.9 Inclusão do servidor "node" no gerenciamento de configurações dos sistemas e serviços utilizando template de configuração já existente': "730",

                    'N3.6.2.8 Configuração de template servidor existente': "729",

                    'N3.6.2.7 Criação (e análise para realização) de contexto de monitoramento': "728",

                    'N3.6.2.6 Configuração sistema de monitoramento': "727",

                    'N3.6.2.5 Configuração de ambiente para novos sistemas': "726",

                    'N3.6.2.4 Configuração de sistemas operacionais servidores Linux (virtual)': "725",

                    'N3.6.2.3 Configuração de serviços de publicação Web - Intranet e Internet': "724",

                    'N3.6.2.2 Configuração de objetos no LDAP': "723",

                    'N3.6.2.1 Configuração de arquivos de aplicações de negócio em servidores de aplicação': "722",

                    'N3.6.1.99 Outras Atividades - Instalar': "721",

                    'N3.6.1.5 Atualização de serviços em sistemas operacionais Linux (por ambiente)': "720",

                    'N3.6.1.4 Atualização de lançamento (release) de versão de sistemas operacionais Linux': "719",

                    'N3.6.1.3 Instalação de sistemas operacionais de servidores Linux (físicos)': "718",

                    'N3.6.1.2 Instalação de serviços para servidores Linux': "717",

                    'N3.6.1.1 Atualização de versão de sistemas operacionais Linux': "716",
                    
                    #Serv. Windows____________________________________________________________________________________________________________

                    'N3.7.3.99 Outras Atividades - Sustentar': "791",

                    'N3.7.3.15 Registro de outras Demandas Rotineiras':"790",
                    
                    'N3.7.3.14 Identificação e análise de problemas (troubleshooting) para a equipe de Suporte à Servidores Windows':"789",
                    
                    'N3.7.3.13 Registro de outros incidentes para a equipe de Suporte à Servidores Windows':"788",
                    
                    'N3.7.3.12 Investigar as vulnerabilidades de SO Windows e propor ações de mitigação':"787",
                    
                    'N3.7.3.11 Servidor – Alerta de memória':"786",
                    
                    'N3.7.3.10 Servidor – Alerta de disco':"785",
                    
                    'N3.7.3.9 Servidor – Alerta de CPU':"784",
                    
                    'N3.7.3.8 Reestabelecer funcionamento de Windows Server devido à indisponibilidade/lentidão do serviço':"783",
                    
                    'N3.7.3.7 Reestabelecer funcionamento de servidor de aplicação devido à indisponibilidade/lentidão do serviço':"782",
                    
                    'N3.7.3.6 Reestabelecer funcionamento do IIS devido à indisponibilidade/lentidão do serviço':"781",
                    
                    'N3.7.3.5 Reestabelecer funcionamento do DNS/WINS (interno) devido à indisponibilidade/lentidão do serviço':"780",
                    
                    'N3.7.3.4 Reestabelecer funcionamento do Active Directory devido à indisponibilidade/lentidão do serviço':"779",
                    
                    'N3.7.3.1 Elaboração de relatório técnico':"778",
                    
                    'N3.7.2.99 Outras Atividades - Configurar':"777",
                    
                    'N3.7.2.19 Criação/Renovação de Certificado':"776",
                    
                    'N3.7.2.18 Resolução de falhas no envio / recebimento de mensagens eletrônicas':"775",
                    
                    'N3.7.2.17 Alteração de serviços em sistemas operacionais Windows':"774",
                    
                    'N3.7.2.16 Movimentação de OU de novos usuários':"773",
                    
                    'N3.7.2.15 Configuração de redirecionamento de mensagens':"772",
                    
                    'N3.7.2.14 Configuração de permissões de acesso à caixas corporativas':"771",
                    
                    'N3.7.2.13 Configuração de listas de distribuição':"770",
                    
                    'N3.7.2.12 Configuração de limites de cotas para envio de e-mail (configurações em mailbox)':"769",
                    
                    'N3.7.2.11 Configuração de VDI':"768",
                    
                    'N3.7.2.10 Configuração de deploy no SCCM':"767",
                    
                    'N3.7.2.9 Configuração de impressora no servidor':"766",
                    
                    'N3.7.2.8 Configuração de ambiente para novos sistemas':"765",
                    
                    'N3.7.2.7 Configuração de ambiente de virtualização':"764",
                    
                    'N3.7.2.6 Configuração de sistemas operacionais servidores Windows (virtual)':"763",
                    
                    'N3.7.2.5 Configuração de serviços de publicação Web - Intranet e Internet':"762",
                    
                    'N3.7.2.4 Configuração de serviço de resolução de nomes - interno':"761",
                    
                    'N3.7.2.3 Configuração de diretivas de grupo':"760",
                    
                    'N3.7.2.2 Configuração de objetos no Active Directory':"759",
                    
                    'N3.7.2.1 Configuração de arquivos de aplicações de negócio em servidores de aplicação':"758",
                    
                    'N3.7.1.99 Outras Atividades - Instalar':"757",
                    
                    'N3.7.1.5 Instalação de site secundário':"756",
                    
                    'N3.7.1.4 Instalação de sistemas operacionais de servidores Windows (físicos)':"755",
                    
                    'N3.7.1.3 Instalação de serviços para servidores Windows':"754",
                    
                    'N3.7.1.2 Instalação de ambiente de virtualização':"753",
                    
                    'N3.7.1.1 Atualização de sistemas operacionais Windows (novas versões ou service packs)':"752",

                    #Segurança_______________________________________________________________________________________________________________

                    'N3.4.3.99 Outras Atividades - Sustentar':"849",
                    
                    'N3.4.3.18 Análise de atividade maliciosa':"848",
                    
                    'N3.4.3.17 Identificação de tráfego malicioso na rede':"847",
                    
                    'N3.4.3.16 Identificação e bloqueio de estação de trabalho comprometida':"846",
                    
                    'N3.4.3.15 Registro de outras Demandas Rotineiras':"845",
                    
                    'N3.4.3.14 Identificação e análise de problemas (troubleshooting) para a equipe de Suporte à Segurança da Informação':"844",
                    
                    'N3.4.3.13 Registro de outros incidentes para a equipe de Suporte à Segurança da Informação':"843",
                    
                    'N3.4.3.12 Investigar as vulnerabilidades de segurança da informação e propor ações de mitigação':"842",
                    
                    'N3.4.3.11 Reestabelecer funcionamento do balanceador de carga devido à indisponibilidade/lentidão do serviço':"841",
                    
                    'N3.4.3.10 Reestabelecer funcionamento do DNS externo devido à indisponibilidade/lentidão do serviço':"840",
                    
                    'N3.4.3.9 Reestabelecer funcionamento do firewall devido à indisponibilidade/lentidão do serviço':"839",
                    
                    'N3.4.3.8 Reestabelecer funcionamento de VPN devido à indisponibilidade/lentidão do serviço':"838",
                    
                    'N3.4.3.7 Verificação da saúde dos dispositivos de segurança do firewall (CPU, Memória, Processos, Sessões)':"837",
                    
                    'N3.4.3.6 Gerar/assinar certificado na CA interna':"836",
                    
                    'N3.4.3.5 Efetuar/restaurar backup de configurações de balanceador de carga':"835",
                    
                    'N3.4.3.4 Efetuar/restaurar backup de configurações de firewall':"834",
                    
                    'N3.4.3.3 Prospecção Técnica de controles de segurança da informação':"833",
                    
                    'N3.4.3.2 Elaboração de políticas de controle de acesso':"832",
                    
                    'N3.4.3.1 Elaboração de relatório técnico':"831",
                    
                    'N3.4.2.99 Outras Atividades - Configurar':"829",
                    
                    'N3.4.2.17 Configuração de roteamento estático':"828",
                    
                    'N3.4.2.16 Configuração de políticas de antivírus':"827",
                    
                    'N3.4.2.15 Criação ou alteração de GSLB':"826",
                    
                    'N3.4.2.14 Criação, alteração ou exclusão de entrada em DNS externo':"825",
                    
                    'N3.4.2.13 Criação ou revogação de certificado digital':"824",
                    
                    'N3.4.2.12 Configuração de VPN':"823",
                    
                    'N3.4.2.11 Configuração de VIP em soluções de segurança':"822",
                    
                    'N3.4.2.10 Configuração de Interface VLAN e Interface Física':"821",
                    
                    'N3.4.2.9 Configuração de profiles de segurança (UTM)':"820",
                    
                    'N3.4.2.8 Configuração de profiles de acesso (Autenticação FSSO)':"819",
                    
                    'N3.4.2.7 Configuração de NAT':"818",
                    
                    'N3.4.2.6 Configuração de Liberação / bloqueio de sites':"817",
                    
                    'N3.4.2.5 Configuração de Entrada no Antispam (Black/White List)':"816",
                    
                    'N3.4.2.4 Configuração de grupos de acesso à internet':"815",
                    
                    'N3.4.2.3 Configuração de balanceador de carga':"814",
                    
                    'N3.4.2.2 Configuração de site no reverso':"813",
                    
                    'N3.4.2.1 Configuração de regra no firewallv':"812",
                    
                    'N3.4.1.99 Outras Atividades - Instalar':"811",
                    
                    'N3.4.1.6 Instalação de agente de antivírus':"810",
                    
                    'N3.4.1.5 Instalação de novo firewall':"809",
                    
                    'N3.4.1.4 Instalação de arquivos de configuração (Restore)':"808",
                    
                    'N3.4.1.3 Atualização de firmware de soluções de segurança':"807",
                    
                    'N3.4.1.2 Atualização de versão do firewall':"806",
                    
                    'N3.4.1.1 Atualização Antispam':"805",
                    
                    #Redes___________________________________________________________________________________________

                    'N3.5.3.99 Outras Atividades - Sustentar':"884",
                    
                    'N3.5.3.18 Registro de outras Demandas Rotineiras':"883",
                    
                    'N3.5.3.17 Identificação e análise de problemas (troubleshooting) para a equipe de Suporte a Rede':"882",
                    
                    'N3.5.3.16 Registro de outros incidentes para a equipe de Suporte à Rede':"881",
                    
                    'N3.5.3.14 Análise de logs':"880",
                    
                    'N3.5.3.13 Investigar as vulnerabilidades de redes e propor ações de mitigação':"879",
                    
                    'N3.5.3.12 Restabelecer funcionamento de VPN devido à lentidão do serviço':"878",
                    
                    'N3.5.3.11 Reestabelecer funcionamento do DHCP devido à indisponibilidade/lentidão do serviço':"877",
                    
                    'N3.5.3.10 Reestabelecer funcionamento da wi-fi devido à indisponibilidade/lentidão do serviço':"876",
                    
                    'N3.5.3.9 Reestabelecer funcionamento da WAN devido à indisponibilidade/lentidão do serviço':"875",
                    
                    'N3.5.3.8 Reestabelecer funcionamento da LAN devido à indisponibilidade/lentidão do serviço':"874",
                    
                    'N3.5.3.7 Reestabelecer funcionamento de conexão à internet devido à indisponibilidade/lentidão do serviço':"873",
                    
                    'N3.5.3.6 Reestabelecer funcionamento de roteador devido à indisponibilidade/lentidão do serviço':"872",
                    
                    'N3.5.3.5 Reestabelecer funcionamento de Access Point devido à indisponibilidade/lentidão do serviço':"871",
                    
                    'N3.5.3.4 Reestabelecer funcionamento de switch devido à indisponibilidade/lentidão do serviço':"870",
                    
                    'N3.5.3.3 Reestabelecer funcionamento de VPN devido à indisponibilidade/lentidão do serviço':"869",
                    
                    'N3.5.3.2 Efetuar/restaurar backup de configurações de Switches / AP / Ativos de redes':"868",
                    
                    'N3.5.3.1 Elaboração de relatório técnico':"866",
                    
                    'N3.5.2.99 Outras Atividades - Configurar':"865",
                    
                    'N3.5.2.12 Prospecção de configuração':"864",
                    
                    'N3.5.2.11 Configuração de usuário em rede sem fio':"863",
                    
                    'N3.5.2.10 Configuração de VPN':"862",
                    
                    'N3.5.2.9 Configuração de políticas de QoS':"861",
                    
                    'N3.5.2.8 Configuração de roteamento (Interface VLAN, rotas estáticas, rotas dinâmicas)':"860",
                    
                    'N3.5.2.7 Configuração de novos access points':"859",

                    'N3.5.2.6 Configuração de ACL':"858",
                    
                    'N3.5.2.5 Configuração de switch na rede perimetral':"857",
                    
                    'N3.5.2.4 Configuração de MAC Adress':"856",
                    
                    'N3.5.2.3 Configuração de VLANs':"855",
                    
                    'N3.5.2.2 Configuração de reservas de endereços IPs':"854",
                    
                    'N3.5.2.1 Configuração de escopos de redes IPs':"853",

                    'N3.5.1.99 Outras Atividades - Instalar':"852",
                    
                    'N3.5.1.2 Instalação física de novos access points':"851",
                    
                    'N3.5.1.1 Instalação física de switchs na rede perimetral':"850",

                    #NOC______________________________________________________
                    
                    'N3.8.3.99 Outras Atividades - Sustentar':"804",
                    
                    'N3.8.3.9 Registro de outras Demandas Rotineiras':"803",
                    
                    'N3.8.3.8 Registro de problemas para a equipe de Suporte à Ambiente de Produção/Monitoramento':"802",

                    'N3.8.3.7 Registro de incidentes para a equipe de Suporte à Ambiente de Produção/Monitoramento':"801",
                    
                    'N3.8.3.6 Acionamento e acompanhamento de fornecedores':"800",
                    
                    'N3.8.3.5 Realizar testes e validações para links de comunicação, serviços / servidores corporativos, e ativos em geral':"799",
                    
                    'N3.8.3.4 Monitoramento de links de comunicação, serviços / servidores corporativos, e ativos em geral':"798",
                    
                    'N3.8.3.3 Executar scripts / jobs para recuperação de um serviço':"797",
                    
                    'N3.8.3.2 Checar (testar) conectividade - Internet e outros (Monitoramento)':"796",

                    'N3.8.3.1 Elaboração de relatório técnico':"795",
                    
                    'N3.8.2.99 Outras Atividades - Configurar':"794",
                    
                    'N3.8.2.1 Criação de imagem de estação de trabalho remota':"793",
                    
                    'N3.8.1.99 Outras Atividades - Instalar':"792",
                 
                    
                }
                return item_dict.get(Item)
            
            Request_Type_Id = Escolha_Do_item(Item)
    
            if self.Querencia_de_dividir.lower() == "sim":
                if self.Limitador_analista >= self.Num_linha01:
                    Analista = self.Analista02

            def criar_chamado(summary, description):
                url = f"{JIRA_URL}/rest/servicedeskapi/request"
                headers = {
                    "Content-Type": "application/json"
                }
                data = {
                    "serviceDeskId": SERVICE_DESK_ID,
                    "requestTypeId": Request_Type_Id,
                    "requestFieldValues": {
                        "summary": summary,
                        "description": description,
                    }
                }
                print(f"Enviando requisição para: {url}")
                print(f"Dados da requisição: {json.dumps(data)}")
                response = requests.post(url, auth=(self.JIRA_USERNAME, self.senha), headers=headers, json=data)
                print(f"Resposta: {response.text}")
                if response.status_code == 201:
                    print("Issue criada com sucesso!")
                    new_issue = response.json()
                    issue_key = new_issue.get('issueKey')
                    return issue_key, new_issue
                else:
                    print(f"Falha ao criar issue. Código de status: {response.status_code}")
                    return None, None

            issue_key, new_issue = criar_chamado(summary, description)
            if issue_key and new_issue:
                print(f"Número da nova issue (CAT): {issue_key}")
                jstr = json.dumps(new_issue)
                j = json.loads(jstr)
                print(f"Objeto JSON completo da issue:\n{j}")
                
                #Add ao analista responsavel
                jira.issue_update(issue_key, fields={"assignee": {"name": Analista}})
                #Add labels
                jira.issue_update(issue_key, fields={'labels': self.labels})
                
                #Tranforma tudo da sting em minusculo e retira os espaços de inicio é de fim
                Querencia_linkar = Querencia_linkar.lower().strip()
                
                #Linka
                if Querencia_linkar == "sim" or  Querencia_linkar == "Sim":
                    jira.create_issue_link({
                        "type": {"id": "10000"},
                        "inwardIssue": {"key": Mud},
                        "outwardIssue": {"key": issue_key}
                    })
                
                #Troca o Enterprise para N3
                url = f"{JIRA_URL}/rest/api/2/issue/{issue_key}"
                payload = {
                    "fields": {
                        "customfield_12901": {"value": novo_valor_enterprise}
                    }
                }
                headers = {
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                response = requests.put(url, json=payload, auth=HTTPBasicAuth(self.JIRA_USERNAME, self.senha), headers=headers)
                #Add comntario
                jira.issue_add_comment(issue_key, comentario)


                #Tranforma tudo da sting em minusculo e retira os espaços de inicio é de fim
                Querencia_de_evidencia = Querencia_de_evidencia.lower().strip()

                #Anexa evidencia 
                if Querencia_de_evidencia == "sim" or Querencia_de_evidencia == "Sim":
                    jira.add_attachment(issue_key, evidencia)
                
                nova_linha = {'Chamado': issue_key}
                df_tabela_chamados = pd.concat([df_tabela_chamados, pd.DataFrame([nova_linha])], ignore_index=True)
                
                df_tabela_chamados.to_excel(self.nome_arquivo, index=False)
                
                
                self.Limitador_analista += 1
                self.c += 1

    
        
        self.image_label.setVisible(True)

    def toggle_password(self, checked):
        if checked:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.eye_button.setIcon(self.eye_icon_open)
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.eye_button.setIcon(self.eye_icon_closed)

    def Encerrar_chamado(self):
        try:
    
            self.nome_arquivo = "chamados_criados.xlsx"
            
            colunas = 'Chamado'

            tabela = pd.read_excel(self.nome_arquivo)
            display(tabela)

            Num_linha = tabela.shape[0]
            linha = 0
            c = 0

            print("Número de linhas no certificado:", Num_linha)

            for linha in range(Num_linha):
                jira = Jira(url=JIRA_URL, username=self.JIRA_USERNAME, password=self.senha)

                issue_key = f"{tabela.loc[linha, colunas]}"
                print(issue_key)

                
                
                
                # Transição de estado para "Encerrado"
                try:
                    jira.issue_transition(issue_key, "Encerrado")
                    print(f"Issue {issue_key} encerrada com sucesso.")
                except Exception as e:
                    print(f"Erro ao encerrar a issue {issue_key}: {e}")

                tabela = tabela.drop(index=c)
                tabela.to_excel(self.nome_arquivo, index=False)

                linha += 1
                c += 1
        
        except Exception as e:
            # Tratar erros gerais para evitar que o software feche
            print(f"Ocorreu um erro ao tentar encerrar os chamados: {e}")
            QMessageBox.critical(self, 'Erro', f"Ocorreu um erro ao encerrar os chamados: {e}")
        



def main():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()