namespace ControlePRJ
{
    partial class RespDSOL
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.bt_Concluido = new System.Windows.Forms.Button();
            this.clb_Analistas = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // bt_Concluido
            // 
            this.bt_Concluido.Location = new System.Drawing.Point(107, 362);
            this.bt_Concluido.Name = "bt_Concluido";
            this.bt_Concluido.Size = new System.Drawing.Size(175, 23);
            this.bt_Concluido.TabIndex = 1;
            this.bt_Concluido.Text = "Concluido";
            this.bt_Concluido.UseVisualStyleBackColor = true;
            this.bt_Concluido.Click += new System.EventHandler(this.bt_Concluido_Click);
            // 
            // clb_Analistas
            // 
            this.clb_Analistas.BackColor = System.Drawing.SystemColors.Info;
            this.clb_Analistas.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.clb_Analistas.CheckOnClick = true;
            this.clb_Analistas.FormattingEnabled = true;
            this.clb_Analistas.Location = new System.Drawing.Point(12, 12);
            this.clb_Analistas.MultiColumn = true;
            this.clb_Analistas.Name = "clb_Analistas";
            this.clb_Analistas.Size = new System.Drawing.Size(380, 330);
            this.clb_Analistas.TabIndex = 2;
            // 
            // RespDSOL
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(407, 397);
            this.Controls.Add(this.clb_Analistas);
            this.Controls.Add(this.bt_Concluido);
            this.HelpButton = true;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(423, 360);
            this.Name = "RespDSOL";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Responsável DSOL";
            this.Load += new System.EventHandler(this.RespDSOL_Load);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.RespDSOL_HelpRequested);            

        }

        #endregion

        private System.Windows.Forms.Button bt_Concluido;
        private System.Windows.Forms.CheckedListBox clb_Analistas;
    }
}