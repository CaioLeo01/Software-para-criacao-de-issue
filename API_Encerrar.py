
import pandas as pd
from atlassian import Jira
from IPython.display import display
from cryptography.fernet import Fernet
from getpass import getpass





Arquivo_ler = r"C:\Users\Caio Manoel\Desktop\Automatização\chamados.xlsx"

comentario = """

Demanda concluída com sucesso

    """

colunas = 'Chamado'
JIRA_URL = "https://jira.itamaraty.gov.br"
JIRA_USERNAME = "caio.leonardo"

# Inicializa um DataFrame vazio para armazenar as issue keys
df_tabela_chamados = pd.DataFrame(columns=['Chamado'])

tabela = pd.read_excel(Arquivo_ler)
display(tabela)


Num_linha = tabela.shape[0]
Num_linha -=1
linha = 0
c = 0

print("Número de linhas no certificado:", Num_linha)





for linha in range(Num_linha + 1):

    jira = Jira(url=JIRA_URL, username=JIRA_USERNAME, password=senha)


    issue_key = f"{tabela.loc[linha, colunas]}"
    print(issue_key)

    
    jira.issue_add_comment(issue_key,comentario )
    jira.issue_transition(issue_key,"Encerrado")


    tabela = tabela.drop(index=c)
    tabela.to_excel(Arquivo_ler, index=False)

    linha+= 1
    c += 1