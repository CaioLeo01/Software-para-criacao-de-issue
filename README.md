Neste software, realizamos a criação de issues no ITILSM Jira através da API do Jira. Para criar as issues, é necessário seguir as etapas abaixo:

Leitura do Arquivo XLSX: O software lê um arquivo .xlsx que contém as informações necessárias para a criação das issues, verificando se ele atende às condições definidas para o processo.

Login do Analista: O analista deve inserir seu login e senha no sistema.

Validação do Usuário: Após o login, o sistema valida o usuário para confirmar que as credenciais estão corretas.

Inserção de Labels: O analista insere as labels (rótulos) que descrevem brevemente a issue que será criada.

Divisão de Responsabilidade (Opcional): O analista tem a opção de dividir a criação das issues com outro usuário. Caso escolha essa opção, deverá informar o login do colaborador com quem compartilhará as tarefas.

Preview de Confirmação: Antes de iniciar a criação das issues, o sistema apresenta um preview para o analista revisar e confirmar as informações:

Nome do arquivo XLSX.
Labels inseridas.
Indicação de divisão de tarefas (se houver).
Início da Criação: Após a confirmação do preview, o analista pressiona o botão "Start", e o sistema inicia o processo de criação das issues utilizando o método POST da API do Jira.

Funcionalidades do Software:
Criação de Issues: As issues são criadas com base nas informações extraídas do arquivo XLSX, associadas ao analista responsável.

Atribuição Automática: As issues criadas são automaticamente atribuídas ao analista que realizou o login.

Anexos e Comentários: O software permite anexar comentários, evidências e labels às issues criadas.

Seleção do Item de Catálogo: As issues são abertas com o item de catálogo apropriado já selecionado.

Encerramento de Chamados: Além de criar os chamados, o sistema permite encerrá-los ao pressionar o botão "Fechar Chamados".
