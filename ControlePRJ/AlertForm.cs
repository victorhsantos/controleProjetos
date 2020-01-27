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
    public partial class AlertForm : Form
    {

        private string alertMensagem;

        public AlertForm(string alertMensagem)
        {
            InitializeComponent();
            this.alertMensagem = alertMensagem;
            lb_Alert.Text = this.alertMensagem;
        }

        private void bt_CloseAlert_Click(object sender, EventArgs e)
        {
            this.Close();
        }        
    }
}
