namespace ControlePRJ
{
    partial class ErrorForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorForm));
            this.lb_Erro = new System.Windows.Forms.Label();
            this.bt_MoreDetailsErro = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tb_ErroMsg = new System.Windows.Forms.RichTextBox();
            this.bt_LessDetailsErro = new System.Windows.Forms.Button();
            this.bt_CloseErro = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // lb_Erro
            // 
            this.lb_Erro.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_Erro.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lb_Erro.Location = new System.Drawing.Point(118, 23);
            this.lb_Erro.Name = "lb_Erro";
            this.lb_Erro.Size = new System.Drawing.Size(440, 68);
            this.lb_Erro.TabIndex = 0;
            this.lb_Erro.Text = "Ocorreu um erro inesperado, se o problema persistir, entre em contato com o Admin" +
    "istrador.";
            // 
            // bt_MoreDetailsErro
            // 
            this.bt_MoreDetailsErro.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_MoreDetailsErro.Image = ((System.Drawing.Image)(resources.GetObject("bt_MoreDetailsErro.Image")));
            this.bt_MoreDetailsErro.Location = new System.Drawing.Point(470, 94);
            this.bt_MoreDetailsErro.Name = "bt_MoreDetailsErro";
            this.bt_MoreDetailsErro.Size = new System.Drawing.Size(28, 25);
            this.bt_MoreDetailsErro.TabIndex = 1;
            this.bt_MoreDetailsErro.UseVisualStyleBackColor = true;
            this.bt_MoreDetailsErro.Click += new System.EventHandler(this.bt_MoreDetailsErro_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(6, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(106, 107);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // tb_ErroMsg
            // 
            this.tb_ErroMsg.BackColor = System.Drawing.SystemColors.Control;
            this.tb_ErroMsg.Location = new System.Drawing.Point(6, 126);
            this.tb_ErroMsg.MaxLength = 500;
            this.tb_ErroMsg.Name = "tb_ErroMsg";
            this.tb_ErroMsg.ReadOnly = true;
            this.tb_ErroMsg.Size = new System.Drawing.Size(552, 64);
            this.tb_ErroMsg.TabIndex = 3;
            this.tb_ErroMsg.Text = "";
            // 
            // bt_LessDetailsErro
            // 
            this.bt_LessDetailsErro.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_LessDetailsErro.Image = ((System.Drawing.Image)(resources.GetObject("bt_LessDetailsErro.Image")));
            this.bt_LessDetailsErro.Location = new System.Drawing.Point(470, 94);
            this.bt_LessDetailsErro.Name = "bt_LessDetailsErro";
            this.bt_LessDetailsErro.Size = new System.Drawing.Size(28, 25);
            this.bt_LessDetailsErro.TabIndex = 4;
            this.bt_LessDetailsErro.UseVisualStyleBackColor = true;
            this.bt_LessDetailsErro.Click += new System.EventHandler(this.bt_LessDetailsErro_Click);
            // 
            // bt_CloseErro
            // 
            this.bt_CloseErro.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_CloseErro.ForeColor = System.Drawing.Color.DarkRed;
            this.bt_CloseErro.Location = new System.Drawing.Point(504, 94);
            this.bt_CloseErro.Name = "bt_CloseErro";
            this.bt_CloseErro.Size = new System.Drawing.Size(54, 25);
            this.bt_CloseErro.TabIndex = 5;
            this.bt_CloseErro.Text = "Fechar";
            this.bt_CloseErro.UseVisualStyleBackColor = true;
            this.bt_CloseErro.Click += new System.EventHandler(this.bt_CloseErro_Click);
            // 
            // ErrorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 125);
            this.Controls.Add(this.bt_CloseErro);
            this.Controls.Add(this.bt_MoreDetailsErro);
            this.Controls.Add(this.bt_LessDetailsErro);
            this.Controls.Add(this.tb_ErroMsg);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lb_Erro);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Erro!";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lb_Erro;
        private System.Windows.Forms.Button bt_MoreDetailsErro;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.RichTextBox tb_ErroMsg;
        private System.Windows.Forms.Button bt_LessDetailsErro;
        private System.Windows.Forms.Button bt_CloseErro;
    }
}