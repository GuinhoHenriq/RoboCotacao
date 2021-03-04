<img src="https://img.shields.io/static/v1?label=Version&message=1.0&color=00B0D8&style=for-the-badge&logo=Probot"/><a href="https://www.selenium.dev/"> <img src="https://img.shields.io/static/v1?label=Selenium&message=3.150.1&color=43B02A&style=for-the-badge&logo=Selenium"/></a> <img src="https://img.shields.io/static/v1?label=SQL&message=SQL Server&color=CC2927&style=for-the-badge&logo=Microsoft SQL Server"/>


<br>
# RoboCotacaoBradesco

## 🚀 Programa de automação que gera cotação de renovação de autos e envia para o e-mail do cliente!

<!--ts-->
   * [Sobre](#Sobre)
   * [Tecnologias](#tecnologias)
   * [Como usar](#como-usar)
      * [Pre Requisitos](#pre-requisitos)
      * [Funcionalidade](#funcionalidade)
<!--te-->

📄 Sobre
=========

Aplicativo Desktop no qual o robô foi desenvolvido para otimizar o processo de gerar e enviar uma prévia de cotação de renovação de seguro de autos para o cliente, eliminando assim o uso de mão de obra humana para efetuar a tarefa,
fazendo assim com que a mão de obra humana seja utilizada em uma outra areá no qual não é efetuado um processo repetitivo.

🛠 Tecnologias
=========

As seguintes tecnologias foram utilizadas para criação do projeto.
  <p> &nbsp • C# - ASP.NET </p>
  <p> &nbsp • Selenium - Web Driver (https://www.selenium.dev/) </p>

📄 Como usar
=========

📜 Pre Requisitos
--------------

Deve-se ter instalado na máquina onde a aplicação ira rodar as seguintes aplicações:
  <p> &nbsp • Google Chrome - <strong>Versão mais atual.</strong> </p>
  <p> &nbsp • Selenium: Chrome Driver <strong>Versão mais atual.</strong> </p>
  <br>
Deve-se ter um banco de dados onde a aplicação ira buscar as informações do cliente:
  <p> &nbsp • SQL Server, Oracle, etc. </p>
  <br>
O Cliente que será trabalhado deve possuir os seguintes requisitos cadastrados no BD
  <p> &nbsp • Nome </p>
  <p> &nbsp • Número de contato - <strong>Com o DDD </strong> </p>
  <p> &nbsp • E-mail de contato </p>
  <p> &nbsp • CEP </p>
  <p> &nbsp • Matricula </p>
  <p> &nbsp • CPF / CNPJ </p>
  <p> &nbsp • Conta corrente </p>
  <p> &nbsp • Agencia  </p>
  <p> &nbsp • Sucursal </p>
  <p> &nbsp • Apolice </p>
  <p> &nbsp • Fim de vigencia próximo da data de execução do rôbo</p>
  
  
💻 Funcionalidade
--------------

<p>Ao iniciar a aplicação o rôbo efetua uma consulta no banco de dados para verificar quais são os cliente com informações aptas para serem trabalhadas no portal do Bradesco</p>
<p>Após verificar se o cliente esta apto o mesmo retorna uma lista com o máximo de 100 registros para que não haja gargalo na aplicação.</p>
<p>Após o retorno da lista o chrome driver entra em ação, abrindo o navegador e preenchendo os campos indicados de forma automatica.</p>
<p>Assim que preenchido todos os campos obrigatorios é gerado um PDF através do portal de vendas do Bradesco e em seguida a aplicação retira este arquivo da aba downloads e salva em um servidor compartilhado
onde o BD tenha acesso, pois assim que salvo neste servidor o BD pega o arquivo e envia por e-mail no endereço encontrado no BD.</p>
<br>
<h3 align="center"> 
	🤖  RoboCotacaoBradesco 🤖 Completo - Em Produção! 💻
</h3>

### ✔️ Features

- [x] Requisição para o BD
- [x] Login no Portal de forma automática
- [x] Preenchimento de campos obrigatorios de forma automática
- [x] Gerar PDF e mover para um local no qual o servidor consiga acessar
- [x] Envio de E-mail para o cliente
