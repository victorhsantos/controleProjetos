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
    public partial class ConcluidoForm : Form
    {
        private string concluidoMensagem;

        public ConcluidoForm(string concluidoMenssagem)
        {
            InitializeComponent();
            this.concluidoMensagem = concluidoMenssagem;
            lb_Concluido.Text = this.concluidoMensagem;
        }

        private void bt_Ok_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
