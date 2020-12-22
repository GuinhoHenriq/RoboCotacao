using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.Events;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Diagnostics;
using System.Timers;
using System.Drawing;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using System.Collections;
using Sikuli4Net.sikuli_JSON;
using Sikuli4Net.sikuli_REST;
using Sikuli4Net.sikuli_UTIL;
using System.IO;

namespace RoboCotacaoBradesco
{
    public partial class RoboCotacao : Form
    {
        #region Variaveis que recebem valor da Base

        string Nome, DDD, numCont, tipoSolic, Campanha, veiculo, agencia, CPFCNPJ, Cia, Sucursal, apolice;
        string itemApolice, sinistro, qtdSinistro, chassi, placa, email, CEP, matricula, Banco, Worksite;

        #endregion

        public RoboCotacao()
        {
            InitializeComponent();
        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            int count = 1;
            int qntdm = 30;

            while (count <= qntdm) ;
            {
                count++;
                loginPortalVendas();
            }
                
            
        }

        #region Metodo loginPortalVendas
        /*
         * Autor:       Guilherme Henrique - 15/12/2020
         * Obs.:        Metodo que efetua o login no portal de vendas do Bradesco
         *              para efetuar a cotação e enviar PDF via e-mail.
         */
         #endregion

        private void loginPortalVendas()
        {
            #region Variaveis IWebElement

            // Variaveis - página de login
            IWebElement user = null;
            IWebElement pass = null;
            IWebElement btnLogin = null;

            // Variaveis - página do Portal de Vendas
            IWebElement btnSeguro = null;

            // Variaveis - página Seguro de automóvel
            IWebElement btnCalculo = null;
            IList<IWebElement> button = null;


            // Variaveis - página Cálculo de Seguro Auto

            // Aba de Solicitante
            IWebElement txtNomeCont = null;
            IWebElement txtDDD = null;
            IWebElement txtTelCont = null;
            IWebElement dpSolicitante = null;
            IWebElement txtCPFCNPJSol = null;
            IWebElement btnFecharAlert = null;
            IWebElement dpCia = null;
            IWebElement txtSinistro = null;
            IWebElement dpBonusAtual = null;
            IWebElement dpBonusAnterior = null;
            IWebElement btnProximo = null;
            IWebElement txtApoSuc = null;
            IWebElement txtApolice = null;
            IWebElement txtApoItem = null;
            IWebElement btnReiniciar = null;

            // TIPO DE SOLICITANTE 6 = FUNCIONÁRIO APOSENTADO

            IWebElement dpSucursal = null;

            // TIPO DE SOLICITANTE 7 = FUNCIONÁRIO

            IWebElement txtMatricula = null;

            // TIPO DE SOLICITANTE 8 = CORRENTISTA

            IWebElement dpCampanha = null;
            IWebElement txtAgencia = null;
            IWebElement dpVeiculo = null;

            // TIPO DE SOLICITANTE 10 = NÃO CORRENTISTA

            IWebElement txtCEPPern = null;

            // TIPO DE SOLICITANTE 20 = WORKSITE 

            IWebElement dpWorksite = null;

            #endregion

            #region Variaveis Fixas

            string URL = "https://wwws.bradescoseguros.com.br/CVCR-PortalDeVendas/acesso/login.do";
            string usuario = "36350597818";
            string senha = "Br832010";

            #endregion

            #region Variaveis Dinâmicas

            string Nome = "EMILIA HORA DA SILVA";
            string DDD = "71";
            string numCont = "999966309";
            string tipoSolic = "FUNCIONÁRIO";
            string veiculo = "PASSEIO";
            string CPFCNPJ = "48973564587";
            string Cia = "BRADESCO SEGUROS S/A";
            string Sucursal = "861";
            string apolice = "051806";
            string itemApolice = "1";
            string sinistro = "Nao";
            string qtdSinistro = "3";
            int textqtdsinistro = Int32.Parse(qtdSinistro);
            string email = "guinho0010@hotmail.com";
            string CEP = "08553030";
            string matricula = "5015626";
            string Banco = "237";
            string Worksite = "AZUL LINHAS AEREAS BRASILEIRAS S/A";
            int classe;
            string bonusAtual;
            string agencia = "0407";

            #endregion

            #region Instanciamento das classes

            OpenQA.Selenium.Chrome.ChromeDriver chromeDriver = null;
            ChromeOptions optionsChrome = new ChromeOptions();
            var chromeDriverService = ChromeDriverService.CreateDefaultService(RoboCotacaoBradesco.Properties.Settings.Default.CaminhoDriverChrome.ToString());
            chromeDriverService.HideCommandPromptWindow = true;

            #endregion

            #region Opções da janela do navegador

            //Adicionando argumentos "SandBox off", "inicio em tela Cheia", "ignorar erros de certificados", "desabilitar popUP de bloqueio"
            optionsChrome.AddArguments("--no-default-browser-check", "--disable-infobars", "no-sandbox", "--start-maximized", "--ignore-certificate-errors", "--disable-popup-blocking", "--app=");
            optionsChrome.AddUserProfilePreference("credentials_enable_service", false);
            optionsChrome.AddUserProfilePreference("profile.password_manager_enabled", false);
            optionsChrome.AddAdditionalCapability("useAutomationExtension", false);

            #endregion

            #region Instanciando opções do navegador e abrindo navegador na URL selecionada

            // Iniciando a instância com as opções previamente configuradas
            chromeDriver = new ChromeDriver(chromeDriverService, optionsChrome);

            //Acessando a URL que está setada fixamente.
            chromeDriver.Navigate().GoToUrl(URL);

            #endregion

            #region Página de login

            // Inserindo o login no campo de usuário
            user = chromeDriver.FindElementById("login");
            user.SendKeys(usuario);

            // Inserindo o Password no campo de senha
            pass = chromeDriver.FindElementById("senha");
            pass.SendKeys(senha);

            // Clicando no Botão de Login
            btnLogin = chromeDriver.FindElementByClassName("btn");
            btnLogin.Click();

            listBox1.Items.Add("Login efetuado com sucesso!");
            listBox1.Items.Add("");
            #endregion

            #region Página seguro de automóvel

            // Alimentando a lista com a quantidade de atributos do tipo "Button" que estão
            // sendo exibidos na tela

            //Tentativa 1 de recuperar as informações na tela e inserir na lista
            button = chromeDriver.FindElements(By.XPath("//button[contains(@name, 'button')]")); Thread.Sleep(2000);

            // Tentativa 2 de recuperar as informações na tela e inserir na lista
            if (button.Count == 0)
            {
                button = chromeDriver.FindElements(By.XPath("//button[contains(@name, 'button')]")); Thread.Sleep(2000);
            }
            //Tentativa 3 de recuperar as informações na tela e inserir na lista
            else if (button.Count == 0)
            {
                button = chromeDriver.FindElements(By.XPath("//button[contains(@name, 'button')]")); Thread.Sleep(2000);
            }

            // Buscando na lista o atributo button que contenha "Seguro de Automóvel"
            btnSeguro = button.FirstOrDefault(x => x.Text.ToLower() == ("Seguro de Automóvel").ToLower()); Thread.Sleep(3000);
            btnSeguro.Click(); Thread.Sleep(5000);

            #endregion

            #region Selecionando o botão de calculo

            // Alterando a página no qual o robo deve procurar os elementos
            chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles[1]); // Página Seguro de Automóvel

            // Clicando no botão de calculo
            btnCalculo = chromeDriver.FindElementByClassName("icon-nav-editar");
            btnCalculo.Click();

            #endregion

            #region Página Cálculo de Seguro Auto - Aba Solicitante

                chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles[2]);

                txtNomeCont = chromeDriver.FindElementById("absc_txtNmContato"); Thread.Sleep(3000);
                txtNomeCont.SendKeys(Nome);

                txtDDD = chromeDriver.FindElementById("absc_txtDddFoneContato"); Thread.Sleep(1500);
                txtDDD.SendKeys(DDD);

                txtTelCont = chromeDriver.FindElementById("absc_txtFoneContato");
                txtTelCont.SendKeys(numCont);

                #endregion

            #region Variaveis DropdownList / Combobox

                dpSucursal = chromeDriver.FindElementById("absc_cmbsucursal");
                var selectSucursal = new SelectElement(dpSucursal);

                dpCampanha = chromeDriver.FindElementById("absc_cmbCampanha");
                var selectCampanha = new SelectElement(dpCampanha);

                dpSolicitante = chromeDriver.FindElementById("absc_cmbTipoSolicitante");
                var selectElement = new SelectElement(dpSolicitante);

                dpVeiculo = chromeDriver.FindElementById("absc_cmbTipoVeiculo");
                var selectVeiculo = new SelectElement(dpVeiculo);

                dpCia = chromeDriver.FindElementById("absegc_cmbCia");
                var selectCia = new SelectElement(dpCia);

                dpBonusAtual = chromeDriver.FindElementById("absegc_cmbClasseBonusAtual");
                var selectBonus = new SelectElement(dpBonusAtual);

                dpWorksite = chromeDriver.FindElementById("absc_cmbworksite");
                var selectWorksite = new SelectElement(dpWorksite);


                #endregion

            #region switch Tipo de Solicitante
                
                    switch (tipoSolic)
                    {
                        #region FUNCIONÁRIO APOSENTADO

                        case "FUNCIONÁRIO APOSENTADO":

                            #region ABA SOLICITANTE

                            selectElement.SelectByText("FUNCIONÁRIO APOSENTADO");

                            selectSucursal.SelectByText("579");

                            txtNomeCont = chromeDriver.FindElementById("absc_txtNmFuncionario");
                            txtNomeCont.SendKeys(Nome);

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA SEGURO

                            selectCia.SelectByText(Cia); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(Sucursal);

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(apolice);

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(itemApolice);


                            SendKeys.Send("{ENTER}");

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {
                               
                            }
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[7]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }

                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro");
                            txtSinistro.SendKeys(sinistro);

                            dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                            dpBonusAnterior.GetAttribute("value");

                            classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) + 1;
                            bonusAtual = "Classe " + classe;

                            if (bonusAtual == "Classe 11")
                            {
                                selectBonus.SelectByText("Classe 10");
                            }

                            else
                            {
                                selectBonus.SelectByText(bonusAtual);
                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA AUTOMÓVEL

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA CALCULAR

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                            btnReiniciar.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + CPFCNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            return;
                            }
                            catch
                            {

                            }

                            try
                            {

                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();
                            
                            }
                            catch
                            {

                            }

                            #endregion

                            break;

                        #endregion

                        #region FUNCIONÁRIO

                        case "FUNCIONÁRIO":

                            #region ABA SOLICITANTE 

                            selectElement.SelectByText("FUNCIONÁRIO");

                            try
                            {
                                txtMatricula = chromeDriver.FindElementById("absc_txtMatricula");
                                txtMatricula.SendKeys(matricula); Thread.Sleep(1000);
                            }
                            catch
                            {
                                listBox1.Items.Add("CPF: " + CPFCNPJ + " Encontra-se sem mátricula");
                                return;
                            }
                            SendKeys.Send("{ENTER}"); Thread.Sleep(2000);

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(4000);
                                btnFecharAlert.Click(); Thread.Sleep(2000);
                            }
                            catch
                            {

                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click(); Thread.Sleep(2000);

                            #endregion

                            #region ABA SEGURO

                            selectCia.SelectByText(Cia);Thread.Sleep(1500);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(Sucursal);

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(apolice);

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(itemApolice);

                            SendKeys.Send("{ENTER}");Thread.Sleep(2000);

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[7]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }

                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro");
                            txtSinistro.SendKeys(sinistro);

                            dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                            dpBonusAnterior.GetAttribute("value");

                            classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) + 1;
                            bonusAtual = "Classe " + classe;

                            if (bonusAtual == "Classe 11")
                            {
                                selectBonus.SelectByText("Classe 10");
                            }

                            else
                            {
                                selectBonus.SelectByText(bonusAtual);
                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA AUTOMÓVEL

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA CALCULAR

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador e reinicia o metodo 
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                            btnReiniciar.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + CPFCNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            return;
                            }
                            catch
                            {

                            }

                            try
                            {

                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();
                            
                            }
                            catch
                            {

                            }

                           
                            #endregion

                            break;

                        #endregion

                        #region CORRENTISTA

                        case "CORRENTISTA":

                            #region ABA SOLICITANTE

                            selectElement.SelectByText("CORRENTISTA");

                            selectCampanha.SelectByText("Sem campanha");

                            txtAgencia = chromeDriver.FindElementById("absc_txtAgencia");
                            txtAgencia.SendKeys(agencia);

                            if (veiculo == "ESPORTIVO" || veiculo == "MOTOCICLETA" || veiculo == "PASSEIO" || veiculo == "PLACA LEVE" || veiculo == "PLACA PESADA PESSOA" || veiculo == "PLACA PESADA CARGA")
                            {

                                selectVeiculo.SelectByText(veiculo); Thread.Sleep(1500);

                                txtCPFCNPJSol = chromeDriver.FindElementById("absc_txtCpfCnpjSol");
                                txtCPFCNPJSol.SendKeys(CPFCNPJ);

                                SendKeys.Send("{ENTER}"); Thread.Sleep(2000);

                                try
                                {
                                    btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                                    btnFecharAlert.Click();
                                }
                                catch
                                {

                                }

                                btnProximo = chromeDriver.FindElementByClassName("btnProximo");
                                btnProximo.Click();

                            }
                            else
                            {

                                selectVeiculo.SelectByText(veiculo); Thread.Sleep(3000);

                                btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                                btnProximo.Click();
                            }

                            #endregion

                            #region ABA SEGURO

                            selectCia.SelectByText(Cia); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(Sucursal);

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(apolice);

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(itemApolice);

                            SendKeys.Send("{ENTER}"); Thread.Sleep(2000);

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[7]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            
                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro");
                            txtSinistro.SendKeys(sinistro);

                            dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                            dpBonusAnterior.GetAttribute("value");

                            classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) + 1;
                            bonusAtual = "Classe " + classe;

                            if (bonusAtual == "Classe 11")
                            {
                                selectBonus.SelectByText("Classe 10");
                            }

                            else
                            {
                                selectBonus.SelectByText(bonusAtual);
                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA AUTOMÓVEL

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA CALCULAR

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador e reinicia o metodo 
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                            btnReiniciar.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + CPFCNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            return;
                            }
                            catch
                            {

                            }

                            try
                            {

                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();
                            
                            }
                            catch
                            {

                            }

                           
                            #endregion

                            break;
                        #endregion

                        #region NÃO CORRENTISTA

                        case "NÃO CORRENTISTA":

                            #region ABA SOLICITANTE 

                            selectElement.SelectByText("NÃO CORRENTISTA");

                            txtCPFCNPJSol = chromeDriver.FindElementById("absc_txtCpfCnpjSeg"); Thread.Sleep(1500);
                            txtCPFCNPJSol.SendKeys(CPFCNPJ);

                            txtCEPPern = chromeDriver.FindElementById("absegc_txtCepPernoiteSolic");
                            txtCEPPern.SendKeys(CEP);

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion
                            
                            #region ABA SEGURO 

                            selectCia.SelectByText(Cia); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol"); Thread.Sleep(2000);
                            txtApoSuc.SendKeys(Sucursal);

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(apolice);

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(itemApolice);

                            SendKeys.Send("{ENTER}");

                             try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[7]/div/div/div[1]/button"); Thread.Sleep(1500);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }

                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro");
                            txtSinistro.SendKeys(sinistro);

                            dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                            dpBonusAnterior.GetAttribute("value");

                            classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) + 1;
                            bonusAtual = "Classe " + classe;

                            if (bonusAtual == "Classe 11")
                            {
                                selectBonus.SelectByText("Classe 10");
                            }

                            else
                            {
                                selectBonus.SelectByText(bonusAtual);
                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA CALCULAR

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador e reinicia o metodo 
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                            btnReiniciar.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + CPFCNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            return;
                            }
                            catch
                            {

                            }

                            try
                            {

                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();
                            
                            }
                            catch
                            {

                            }

                            #endregion

                            break;
                        #endregion

                        #region WORKSTIE

                        case "WORKSITE":

                            // ABA SOLICITANTE

                            selectElement.SelectByText("WORKSITE");

                            selectWorksite.SelectByText(Worksite);

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            // ABA SEGURO

                            selectCia.SelectByText(Cia);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(Sucursal);

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(apolice);

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(itemApolice);

                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro");
                            txtSinistro.SendKeys(sinistro);

                            if (sinistro == "SIM")
                            {

                                txtQtdSinistro = chromeDriver.FindElementById("absegc_txtQtdSinistro");
                                txtQtdSinistro.SendKeys(qtdSinistro);

                                dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                                dpBonusAnterior.GetAttribute("value"); Thread.Sleep(3000);

                                classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) - textqtdsinistro;
                                bonusAtual = "Classe " + classe;

                                selectBonus.SelectByText(bonusAtual);

                                btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                                btnProximo.Click();

                            }
                            else
                            {

                                dpBonusAnterior = chromeDriver.FindElementById("absegc_cmbClasseBonusAnte");
                                dpBonusAnterior.GetAttribute("value");

                                classe = Int32.Parse(dpBonusAnterior.GetAttribute("value")) + 1;
                                bonusAtual = "Classe " + classe;

                                selectBonus.SelectByText(bonusAtual);

                                btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                                btnProximo.Click();

                            }

                            #region ABA CALCULAR

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador e reinicia o metodo 
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click();
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                            btnReiniciar.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + CPFCNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            return;
                            }
                            catch
                            {

                            }

                            try
                            {

                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();
                            
                            }
                            catch
                            {

                            }
                            #endregion

                            break;
                        #endregion

                    }

            #endregion

            #region Gerar PDF
                    // Click no botão de gerar o PDF

                    btnProximo = chromeDriver.FindElementByClassName("btnGerarPdf"); Thread.Sleep(2000);
                    btnProximo.Click(); Thread.Sleep(2000);

                    btnProximo = chromeDriver.FindElementByXPath("/html/body/div[2]/div/div/div[2]/form/div[3]/button[3]"); Thread.Sleep(3000);
                    btnProximo.Click(); Thread.Sleep(2000);
                    

            // Fechar todos os alertas
                try
                    {
                        btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(2000);
                        btnFecharAlert.Click(); Thread.Sleep(2000);

                        btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[2]/div/div/div[2]/form/div[3]/button[2]");
                        btnFecharAlert.Click(); Thread.Sleep(2000);

                        try
                        {
                            btnFecharAlert = chromeDriver.FindElementById("btn_cancelar_gerarpdf");
                            btnFecharAlert.Click();
                        }
                        catch
                        {

                        }

                    }
                    catch
                    {
                        
                    }
                    
                    

                    try
                    {
                        btnProximo = chromeDriver.FindElementByXPath("/html/body/div[2]/div/div/div[2]/form/div[3]/button[3]"); Thread.Sleep(3000);
                        btnProximo.Click(); Thread.Sleep(2000);
                    }
                    catch
                    {

                    }

                    try
                    {
                        btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(3000);
                        btnFecharAlert.Click();
                    }
                    catch
                    {

                    }

                    try
                    {
                        btnProximo = chromeDriver.FindElementById("btn_fechar_gerarpdf"); Thread.Sleep(2000);
                        btnProximo.Click();
                    }
                    catch
                    {

                    }

                btnReiniciar = chromeDriver.FindElementByClassName("btnRestart");
                btnReiniciar.Click();

                RenameFile("C:/Users/881912/Downloads/demonstrativo.pdf","C:/Users/881912/Downloads/"+ CPFCNPJ + ".pdf");

                listBox1.Items.Add("Cotação para o CPF: " + CPFCNPJ + " OK!");
                listBox1.Items.Add("------------------------------------------------------------------------------------------------");

                chromeDriver.Quit();
            
            #endregion

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void RenameFile(string originalName, string newName)
        {
            try
            {
                File.Move(originalName, newName);
            }
            catch
            {
                return;
            }
        }
    }
}