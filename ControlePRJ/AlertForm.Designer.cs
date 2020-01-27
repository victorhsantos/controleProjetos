namespace ControlePRJ
{
    partial class AlertForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AlertForm));
            this.bt_CloseAlert = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lb_Alert = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // bt_CloseAlert
            // 
            this.bt_CloseAlert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_CloseAlert.ForeColor = System.Drawing.Color.DarkOrange;
            this.bt_CloseAlert.Location = new System.Drawing.Point(505, 91);
            this.bt_CloseAlert.Name = "bt_CloseAlert";
            this.bt_CloseAlert.Size = new System.Drawing.Size(54, 25);
            this.bt_CloseAlert.TabIndex = 10;
            this.bt_CloseAlert.Text = "Ok";
            this.bt_CloseAlert.UseVisualStyleBackColor = true;
            this.bt_CloseAlert.Click += new System.EventHandler(this.bt_CloseAlert_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(7, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(106, 107);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // lb_Alert
            // 
            this.lb_Alert.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_Alert.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lb_Alert.Location = new System.Drawing.Point(119, 26);
            this.lb_Alert.Name = "lb_Alert";
            this.lb_Alert.Size = new System.Drawing.Size(440, 62);
            this.lb_Alert.TabIndex = 6;
            this.lb_Alert.Text = "Atenção!";
            // 
            // AlertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 125);
            this.Controls.Add(this.bt_CloseAlert);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lb_Alert);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AlertForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Atenção!";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button bt_CloseAlert;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lb_Alert;
    }
}