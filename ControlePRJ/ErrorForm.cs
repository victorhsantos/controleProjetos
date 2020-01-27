using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlePRJ
{
    public partial class ErrorForm : Form
    {
        private Exception ex;

        public ErrorForm(Exception ex)
        {
            InitializeComponent();
            this.ex = ex;
            tb_ErroMsg.Text = "Detalhes do erro:\n\n" + this.ex.Message;

            if (ex.Message.Length >= 15)
                if (ex.Message.Substring(0, 15) == "Duplicate entry")
                    lb_Erro.Text = "Este projeto já existe. Verifique o Código do Projeto e tente novamente.";
                else
                    lb_Erro.Text = "Ocorreu um erro inesperado, se o problema persistir, entre em contato com o Administrador.";
        }

        private void bt_MoreDetailsErro_Click(object sender, EventArgs e)
        {
            this.Size = new Size(583, 240);
            bt_MoreDetailsErro.Visible = false;
            bt_LessDetailsErro.Visible = true;
        }

        private void bt_LessDetailsErro_Click(object sender, EventArgs e)
        {
            this.Size = new Size(583, 163);
            bt_LessDetailsErro.Visible = false;
            bt_MoreDetailsErro.Visible = true;
        }

        private void bt_CloseErro_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
