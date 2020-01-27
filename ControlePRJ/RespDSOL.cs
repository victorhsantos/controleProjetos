using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data.MySqlClient;


namespace ControlePRJ
{
    public partial class RespDSOL : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_projeto;uid=admin;server = 192.168.10.6; database = controle_projeto; uid = admin; pwd = accenture");
        private string nomes;

        public string Nomes
        {
            get
            {
                return nomes;
            }
            set
            {
                nomes = value;
            }
        }

        public RespDSOL()
        {
            InitializeComponent();            
        }

        private void RespDSOL_Load(object sender, EventArgs e)
        {
            try
            {                
                if (clb_Analistas.Items.Count == 0)
                {
                    //POVOANDO CHECK LISTBOX DE ANALISTAS
                    bdConn.Open();
                    string query = "select nome from analistas order by nome;";
                    MySqlCommand command = new MySqlCommand(query, bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        clb_Analistas.Items.Add(dr["nome"].ToString());
                    dr.Close();
                    bdConn.Close();
                }

            }
            catch (Exception ex)
            {
                bdConn.Close();
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }            
        }       

        private void bt_Concluido_Click(object sender, EventArgs e)
        {

            if (clb_Analistas.CheckedItems.Count != 0)
            {
                for (int x = 0; x <= clb_Analistas.CheckedItems.Count - 1; x++)
                {
                    nomes += clb_Analistas.CheckedItems[x].ToString();

                    if (x != clb_Analistas.CheckedItems.Count - 1)
                        nomes += ", ";
                }

                this.Close();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Nenhum nome foi selecionado. Deseja sair?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dialogResult == DialogResult.Yes)                
                    this.Close();
            }
        }

        //BOTÃO HELP
        private void RespDSOL_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            MessageBox.Show("Selecionar os nomes dos responsáveis pelo DSOL!","Ajuda!", MessageBoxButtons.OK,MessageBoxIcon.Information);
        }        
    }
}
