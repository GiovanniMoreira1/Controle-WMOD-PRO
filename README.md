## 🛠️ Sistema de Inventário - Sala WinMOD PRO

Este projeto automatiza o controle do fluxo de hardware e demais equipamentos do laboratório **WinMOD PRO**.  
O sistema integra respostas enviadas por formulários externos a um banco de dados relacional, possibilitando a gestão eficiente, auditável e em tempo real de empréstimos, liberações e entradas de equipamentos.

## 🚀 Como funciona?

O ecossistema é composto por dois componentes principais que garantem a integridade dos dados e são independentes entre si:

- **Watcher (Serviço em Background):**  
  Aplicação que monitora continuamente as novas submissões dos formulários Microsoft Forms. Ao receber um pedido, o Watcher valida os dados e executa a operação correspondente diretamente no banco de dados.  

- **Interface Administrativa (Tkinter):**  
  Aplicação desktop destinada aos gestores do laboratório, que permite a visualização do inventário em tempo real, a gestão de equipamentos e a extração de relatórios.

## 📊 Estrutura de Dados

O sistema utiliza um banco de dados relacional com quatro tabelas principais:

- **Equipamentos:** Registro dos itens, incluindo IDs, descrições e status atual.  
- **Funcionários:** Cadastro dos usuários autorizados a solicitar movimentações.  
- **Operações:** Histórico detalhado de entradas e saídas, registrando quem realizou, quando e o quê.  
- **Empréstimos:** Monitoramento das informações referentes aos empréstimos realizados.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python  
- **Interface Gráfica:** Tkinter  
- **Monitoramento:** Script Watcher para integração dos formulários via leitura de e-mails.  
- **Banco de Dados:** PostgreSQL (modelo relacional).
