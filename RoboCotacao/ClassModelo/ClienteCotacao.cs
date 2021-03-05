using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoboCotacaoBradesco.ClassModelo
{
    class ClienteCotacao
    {

        #region Informações da classe
        /*
         * Autor:   Guilherme Henrique - 04/12/2020
         * Obs.:    Essa classe representa o modelo da table #TB_COTACAO (SERVER TMKT_ZL_DB26 atualmente)
         *          Procedure - STP_CARREGA_MAILING_ROBO_COTACAO
                    Os dados dessa table representa os usuários que devem receber uma cotação automática.
         */
        #endregion


        #region Propriedades da Classe

        public int Cliente_SEQ { get; set; }
        public string Cliente_Clicodigo { get; set; }
        public string Cliente_CPF_CNPJ { get; set; }
        public string Cliente_Nome { get; set; }
        public int Cliente_DDD { get; set; }
        public int Cliente_Telefone { get; set; }
        public string Cliente_campanha { get; set; }
        public int Cliente_Conta {get; set;} 
        public int Cliente_Sucursal {get; set;}
        public int Cliente_Apolice {get; set;}
        public int Cliente_Item {get; set;}
        public DateTime Cliente_DtFimVigencia {get; set;}
        public string Cliente_Matricula { get; set; }
        public string Cliente_NomeWorksite { get; set;}
        public int Cliente_CEP { get; set;}
        public int Cliente_Agencia { get; set;}
        public string Cliente_Email { get; set; }


        #endregion 
       
    }
}
