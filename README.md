## 🛠️ Sistema de Inventário - Sala WinMOD PRO
Este projeto automatiza o controle de fluxo de hardware no laboratório WinMOD PRO. O sistema integra respostas de formulários externos com um banco de dados relacional, permitindo a gestão de empréstimos, liberações e entradas de equipamentos de forma auditável e eficiente.

## 🚀 Como funciona?
O ecossistema é dividido em dois componentes principais que garantem a integridade dos dados e a facilidade de uso: <br>
- Watcher (Serviço de Background): Uma aplicação headless que monitora constantemente as novas entradas no Microsoft/Google Forms. Assim que um funcionário submete um pedido, o Watcher valida os dados e executa a operação correspondente no banco de dados. <br>
- Interface Administrativa (Tkinter): Uma aplicação desktop para os gestores do laboratório. Permite a visualização do inventário em tempo real, gestão de usuários e extração de relatórios.

## 📊 Estrutura de Dados
O sistema utiliza um banco de dados relacional estruturado em 4 tabelas principais:
Equipamentos: Registro de itens (IDs, descrição, status atual).
Funcionários: Cadastro de quem pode solicitar movimentações.
Operações: Histórico de entradas e saídas (quem, quando e o quê).
Empréstimos: Tabela que monitora os dados dos empréstimos feitos.

## 🛠️ Tecnologias Utilizadas
Linguagem: Python <br>
GUI: Tkinter <br>
Monitoramento: Script Watcher para integração de Forms via leitura de e-mails. <br>
Banco de Dados: PostgreSQL (Modelo Relacional). <br>
