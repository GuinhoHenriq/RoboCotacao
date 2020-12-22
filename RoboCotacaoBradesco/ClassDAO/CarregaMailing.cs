using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Net;
using System.IO;
using System.Web;
using System.Configuration;
using RoboCotacaoBradesco.ClassModelo;

namespace RoboCotacaoBradesco.ClassDAO
{
    class CarregaMailing
    {
        public List<ClienteCotacao> CarregaMalingCotacao (ClienteCotacao objAcessUser)
        {
            System.Data.SqlClient.SqlConnection conexao = new SqlConnection(RoboCotacaoBradesco.Properties.Settings.Default.AcessoBanco);

            SqlCommand comando = new SqlCommand();
            DataSet ds = new DataSet();
            SqlDataAdapter da;

            comando.Connection = conexao;
            comando.CommandType = CommandType.StoredProcedure;
            comando.CommandText = "STP_CARREGA_MAILING_ROBO_COTACAO";
            comando.CommandTimeout = 3000;
            da = new SqlDataAdapter(comando);
            da.Fill(ds, "Dados");
            conexao.Close();

            List<ClienteCotacao> listaCliente = new List<ClienteCotacao>();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                listaCliente.Add(new ClienteCotacao() { Cliente_CPF_CNPJ = ds.Tables[0].Rows[i]["CPFCNPJ"].ToString() });
                listaCliente.Add(new ClienteCotacao() { Cliente_Nome = ds.Tables[0].Rows[i]["NOME"].ToString() });
                listaCliente.Add(new ClienteCotacao() { Cliente_DDD = Convert.ToInt32(ds.Tables[0].Rows[i]["DDD"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_Telefone = Convert.ToInt32(ds.Tables[0].Rows[i]["TELEFONE"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_campanha = ds.Tables[0].Rows[i]["CAMPANHA"].ToString() });
                listaCliente.Add(new ClienteCotacao() { Cliente_Matricula = ds.Tables[0].Rows[i]["MATRICULA"].ToString() });
                listaCliente.Add(new ClienteCotacao() { Cliente_Agencia = Convert.ToInt32(ds.Tables[0].Rows[i]["AGENCIA"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_Conta = Convert.ToInt32(ds.Tables[0].Rows[i]["CONTA"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_Sucursal = Convert.ToInt32(ds.Tables[0].Rows[i]["SUCURSAL"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_Apolice = Convert.ToInt32(ds.Tables[0].Rows[i]["APOLICE"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_Item = Convert.ToInt32(ds.Tables[0].Rows[i]["MATRICULA"].ToString()) });
                listaCliente.Add(new ClienteCotacao() { Cliente_DtFimVigencia = Convert.ToDateTime(ds.Tables[0].Rows[i]["FIM_VIGENCIA"].ToString()) });

                
            }

            return listaCliente;
        }
    }
}
