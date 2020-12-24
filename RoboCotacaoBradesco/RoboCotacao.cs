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
using System.IO;
using RoboCotacaoBradesco.ClassDAO;
using RoboCotacaoBradesco.ClassModelo;

namespace RoboCotacaoBradesco
{
    public partial class RoboCotacao : Form
    {
        #region Variaveis Globais

        int classe, CliCEP;
        string bonusAtual, Worksite, validaCampo;
        string veiculo = "PASSEIO";
        int contNProcss = 0;
        int contProcss = 0;

        #endregion

         delegate void SimplesDelegate();

        public RoboCotacao()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            
            Thread th = new Thread(new ThreadStart(this.CarregaMalingCotacao));
            th.Start();  Thread.Sleep(3000);
      
        }

        public void msgLogin()
        {
            listBox1.Items.Add("Login efetuado com sucesso!");
        }


        #region Metodo loginPortalVendas
        /*
         * Autor:       Guilherme Henrique - 15/12/2020
         * Obs.:        Metodo que efetua o login no portal de vendas do Bradesco
         *              para efetuar a cotação e enviar PDF via e-mail.
         */
         #endregion

        private void loginPortalVendas(ClienteCotacao objCliente)
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
            IWebElement txtLotacao = null;
            IWebElement txtAnoVeic = null;

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

            #endregion

            #region Instanciamento das classes

            OpenQA.Selenium.Chrome.ChromeDriver chromeDriver = null;
            ChromeOptions optionsChrome = new ChromeOptions();
            var chromeDriverService = ChromeDriverService.CreateDefaultService(RoboCotacaoBradesco.Properties.Settings.Default.CaminhoDriverChrome.ToString());
            chromeDriverService.HideCommandPromptWindow = true;
            SimplesDelegate SimplesDelegate = new SimplesDelegate(msgLogin);
            CarregaMailing Dados = new CarregaMailing();

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
            user.SendKeys(RoboCotacaoBradesco.Properties.Settings.Default.usuario);

            // Inserindo o Password no campo de senha
            pass = chromeDriver.FindElementById("senha");
            pass.SendKeys(RoboCotacaoBradesco.Properties.Settings.Default.senha);

            // Clicando no Botão de Login
            btnLogin = chromeDriver.FindElementByClassName("btn");
            btnLogin.Click();

            SimplesDelegate.Invoke();
            
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

                listBox1.Items.Add("Processando CPF: " + objCliente.Cliente_CPF_CNPJ);

            

                txtNomeCont = chromeDriver.FindElementById("absc_txtNmContato"); Thread.Sleep(3000);
                txtNomeCont.SendKeys(objCliente.Cliente_Nome);

                txtDDD = chromeDriver.FindElementById("absc_txtDddFoneContato"); Thread.Sleep(1500);
                txtDDD.SendKeys(objCliente.Cliente_DDD.ToString());

                txtTelCont = chromeDriver.FindElementById("absc_txtFoneContato");
                txtTelCont.SendKeys(objCliente.Cliente_Telefone.ToString());

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
                
                    switch (objCliente.Cliente_campanha)
                    {
                        #region FUNCIONÁRIO APOSENTADO

                        case "FUNCIONÁRIO APOSENTADO":

                            #region ABA SOLICITANTE

                            selectElement.SelectByText("FUNCIONÁRIO APOSENTADO");

                            selectSucursal.SelectByText("579");

                            txtNomeCont = chromeDriver.FindElementById("absc_txtNmFuncionario");
                            txtNomeCont.SendKeys(objCliente.Cliente_Nome);

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA SEGURO

                            selectCia.SelectByText("BRADESCO SEGUROS S/A"); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(objCliente.Cliente_Sucursal.ToString());

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(objCliente.Cliente_Apolice.ToString());

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(objCliente.Cliente_Item.ToString());


                            SendKeys.SendWait("{ENTER}");

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
                            txtSinistro.SendKeys("Nao");

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

                            try
                            {

                                txtLotacao = chromeDriver.FindElementById("abcc_txtLotacaoVeiculo"); Thread.Sleep(3000);
                                validaCampo = txtLotacao.GetAttribute("value");
                                if (validaCampo == "")
                                {
                                    txtLotacao.SendKeys("0"); Thread.Sleep(2000);
                                }
                            }
                            catch
                            {

                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            // Erro ao tentar calcular - fecha o navegador
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click(); Thread.Sleep(3000);
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

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("");
                            txtNProcessado.Text = (contNProcss++).ToString();
                            
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
                                txtMatricula.SendKeys(objCliente.Cliente_Matricula); Thread.Sleep(1000);
                            }
                            catch
                            {
                            }
                            SendKeys.SendWait("{ENTER}"); Thread.Sleep(2000);

                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(4000);
                                btnFecharAlert.Click(); Thread.Sleep(3000);
                            }
                            catch
                            {

                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

                            #endregion

                            #region ABA SEGURO
                            try
                            {
                                selectCia.SelectByText("BRADESCO SEGUROS S/A"); Thread.Sleep(3000);

                                txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol"); Thread.Sleep(1500);
                                txtApoSuc.SendKeys(objCliente.Cliente_Sucursal.ToString());

                                txtApolice = chromeDriver.FindElementById("absegc_txtCapol"); Thread.Sleep(1500);
                                txtApolice.SendKeys(objCliente.Cliente_Apolice.ToString());

                                txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol"); Thread.Sleep(1500);
                                txtApoItem.SendKeys(objCliente.Cliente_Item.ToString());

                                SendKeys.SendWait("{ENTER}"); Thread.Sleep(2000);

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
                             
                            txtSinistro = chromeDriver.FindElementById("absegc_cmbHouveSinistro"); Thread.Sleep(2000);
                            txtSinistro.SendKeys("Nao"); Thread.Sleep(2000);

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
                    }
                            catch
                            {

                            }
                            #endregion

                            #region ABA AUTOMÓVEL

                            txtAnoVeic = chromeDriver.FindElementById("abac_txtAnoFabricacao");
                            string anoVeic = txtAnoVeic.GetAttribute("value");


                            if (anoVeic == "")
                            {
                                listBox1.Items.Add("Erro ao tentar Processar o CPF: " + objCliente.Cliente_CPF_CNPJ + " ...");
                                listBox1.Items.Add("");
                                txtNProcessado.Text = (txtNProcessado.Text + 1).ToString();
                                chromeDriver.Quit();
                                return;
                                
                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion

                            #region ABA CALCULAR

                            try
                            {

                                txtLotacao = chromeDriver.FindElementById("abcc_txtLotacaoVeiculo"); Thread.Sleep(3000);
                                validaCampo = txtLotacao.GetAttribute("value");
                                if (validaCampo == "")
                                {
                                    txtLotacao.SendKeys("0"); Thread.Sleep(2000);
                                }                               
                            }
                            catch
                            {

                            }

                            try
                            {
                                btnProximo = chromeDriver.FindElementByClassName("btnCalcular"); Thread.Sleep(2000);
                                btnProximo.Click(); Thread.Sleep(2000);
                            }
                            catch
                            {
                                
                            }

                            // Erro ao tentar calcular - fecha o navegador e reinicia o metodo 
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click(); Thread.Sleep(3000);

                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click(); Thread.Sleep(3000);
                                
                            }
                            catch
                            {

                            }
                            
                            try
                            {
                            btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[1]/button");
                            btnFecharAlert.Click();

                            chromeDriver.Quit();

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("------------------------------------------------------------------------------------------------");
                            txtNProcessado.Text = (contNProcss++).ToString();
                            
                                return;
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
                            txtAgencia.SendKeys(objCliente.Cliente_Agencia.ToString());

                            if (veiculo == "ESPORTIVO" || veiculo == "MOTOCICLETA" || veiculo == "PASSEIO" || veiculo == "PLACA LEVE" || veiculo == "PLACA PESADA PESSOA" || veiculo == "PLACA PESADA CARGA")
                            {

                                selectVeiculo.SelectByText("PASSEIO"); Thread.Sleep(1500);

                                txtCPFCNPJSol = chromeDriver.FindElementById("absc_txtCpfCnpjSol");
                                txtCPFCNPJSol.SendKeys(objCliente.Cliente_CPF_CNPJ);

                                SendKeys.SendWait("{ENTER}"); Thread.Sleep(2000);

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

                            selectCia.SelectByText("BRADESCO SEGUROS S/A"); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(objCliente.Cliente_Sucursal.ToString());

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(objCliente.Cliente_Apolice.ToString());

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(objCliente.Cliente_Item.ToString());

                            SendKeys.SendWait("{ENTER}"); Thread.Sleep(2000);

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
                            txtSinistro.SendKeys("Nao");

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

                            try
                            {

                                txtLotacao = chromeDriver.FindElementById("abcc_txtLotacaoVeiculo"); Thread.Sleep(3000);
                                validaCampo = txtLotacao.GetAttribute("value");
                                if (validaCampo == "")
                                {
                                    txtLotacao.SendKeys("0"); Thread.Sleep(2000);
                                }
                            }
                            catch
                            {

                            }

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

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("");
                            txtNProcessado.Text = (contNProcss++).ToString();
                            
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
                            txtCPFCNPJSol.SendKeys(objCliente.Cliente_CPF_CNPJ);

                            txtCEPPern = chromeDriver.FindElementById("absegc_txtCepPernoiteSolic");
                            txtCEPPern.SendKeys(CliCEP.ToString());

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion
                            
                            #region ABA SEGURO 

                            selectCia.SelectByText("BRADESCO SEGUROS S/A"); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol"); Thread.Sleep(2000);
                            txtApoSuc.SendKeys(objCliente.Cliente_Sucursal.ToString());

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(objCliente.Cliente_Apolice.ToString());

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(objCliente.Cliente_Item.ToString());

                            SendKeys.SendWait("{ENTER}");

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
                            txtSinistro.SendKeys("Nao");

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

                            try
                            {

                                txtLotacao = chromeDriver.FindElementById("abcc_txtLotacaoVeiculo"); Thread.Sleep(3000);
                                validaCampo = txtLotacao.GetAttribute("value");
                                if (validaCampo == "")
                                {
                                    txtLotacao.SendKeys("0"); Thread.Sleep(2000);
                                }
                            }
                            catch
                            {

                            }

                            btnProximo = chromeDriver.FindElementByClassName("btnCalcular");Thread.Sleep(2000);
                            btnProximo.Click(); Thread.Sleep(2000);

             
                            try
                            {
                                btnFecharAlert = chromeDriver.FindElementByXPath("/html/body/div[8]/div/div/div[3]/div/button"); Thread.Sleep(3000);
                                btnFecharAlert.Click(); Thread.Sleep(5000);
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

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("");
                            
                            contNProcss++;

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

                            #region ABA SOLICITANTE

                            selectElement.SelectByText("WORKSITE");

                            selectWorksite.SelectByText(Worksite);

                            btnProximo = chromeDriver.FindElementByClassName("btnProximo"); Thread.Sleep(1500);
                            btnProximo.Click();

                            #endregion
                            
                            #region ABA SEGURO

                            selectCia.SelectByText("BRADESCO SEGUROS S/A"); Thread.Sleep(3000);

                            txtApoSuc = chromeDriver.FindElementById("absegc_txtCsucApol");
                            txtApoSuc.SendKeys(objCliente.Cliente_Sucursal.ToString());

                            txtApolice = chromeDriver.FindElementById("absegc_txtCapol");
                            txtApolice.SendKeys(objCliente.Cliente_Apolice.ToString());

                            txtApoItem = chromeDriver.FindElementById("absegc_txtCitemApol");
                            txtApoItem.SendKeys(objCliente.Cliente_Item.ToString());

                            SendKeys.SendWait("{ENTER}");

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
                            txtSinistro.SendKeys("Nao");

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

                            try
                            {

                                txtLotacao = chromeDriver.FindElementById("abcc_txtLotacaoVeiculo"); Thread.Sleep(3000);
                                validaCampo = txtLotacao.GetAttribute("value");
                                if (validaCampo == "")
                                {
                                    txtLotacao.SendKeys("0"); Thread.Sleep(2000);
                                }
                            }
                            catch
                            {

                            }

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

                            listBox1.Items.Add("Erro ao tentar calcular valores do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("");
                            contNProcss++;
                                
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

                    }
                catch
                    {

                    }


                try
                    {
                            btnFecharAlert = chromeDriver.FindElementById("btn_cancelar_gerarpdf");
                            btnFecharAlert.Click();

                            listBox1.Items.Add("Erro ao tentar Extrair PDF do CPF: " + objCliente.Cliente_CPF_CNPJ);
                            listBox1.Items.Add("");

                            txtNProcessado.Text = (contNProcss++).ToString();
                            
                            chromeDriver.Quit();
                            return;
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
                    RenameFile("C:/Users/881912/Downloads/demonstrativo.pdf", "C:/Users/881912/Downloads/Robo/" + objCliente.Cliente_CPF_CNPJ + ".pdf"); Thread.Sleep(4000);
                    Dados.EnviaEmailCli(objCliente);
                

                chromeDriver.Quit();
            
            #endregion

        }
        private void CarregaMalingCotacao()
        {
           
            ClassDAO.CarregaMailing Dados = new ClassDAO.CarregaMailing();
            ClassModelo.ClienteCotacao objAccessUser = new ClassModelo.ClienteCotacao();
            List<ClienteCotacao> listaCliente = new List<ClienteCotacao>();
            DataSet ds = new DataSet();
            

            listaCliente = Dados.CarregaMalingCotacao();

            string qtdRec = listaCliente.Count.ToString();

            txtRecebido.Text = qtdRec;

            foreach (var item in listaCliente)
            {

                progressBar1.Increment(1);
                loginPortalVendas(item);
                
            }
  

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void RenameFile(string originalName, string newName)
        {
            try
            {
                File.Move(originalName, newName); Thread.Sleep(5000);
                txtProcessado.Text = (contProcss++).ToString();
                listBox1.Items.Add("CPF Processado com Sucesso!");
                listBox1.Items.Add("");


                
            }
            catch
            {
                txtNProcessado.Text = (contNProcss++).ToString();
                listBox1.Items.Add("CPF Não Processado");
                listBox1.Items.Add("");
                return;
            }
        }

        private void RoboCotacao_Load(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
        
    }
}