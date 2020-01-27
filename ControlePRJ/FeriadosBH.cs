using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;

namespace FeriadosBH
{
    public class Feriados
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=feriados_bh2017;uid=admin;server = 192.168.10.6; database = feriados_bh2017; uid = admin; pwd = accenture; Allow Zero Datetime=True");

        public Dictionary<DateTime, string> getListaFeriados()
        {
            Dictionary<DateTime, string> listaFeriados = new Dictionary<DateTime,string>();

            try
            {
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT data, descricao FROM feriados", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                while (dr.Read())
                    listaFeriados.Add(DateTime.Parse(dr["data"].ToString()), dr["descricao"].ToString());
                dr.Close();
            }
            catch
            {
                listaFeriados.Clear();
            }
            finally
            {
                bdConn.Close();
            }

            return listaFeriados;
        }               
    }
}
