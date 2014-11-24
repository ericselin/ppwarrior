namespace PowerPoint_Warrior
{
    partial class FormSettings
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
            this.components = new System.ComponentModel.Container();
            this.chkLogging = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbEmail = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.errorpSettings = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorpSettings)).BeginInit();
            this.SuspendLayout();
            // 
            // chkLogging
            // 
            this.chkLogging.AutoSize = true;
            this.chkLogging.Checked = true;
            this.chkLogging.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLogging.Enabled = false;
            this.chkLogging.Location = new System.Drawing.Point(18, 227);
            this.chkLogging.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkLogging.Name = "chkLogging";
            this.chkLogging.Size = new System.Drawing.Size(315, 24);
            this.chkLogging.TabIndex = 0;
            this.chkLogging.Text = "Participate in anonymous usage logging";
            this.chkLogging.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(14, 260);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(369, 76);
            this.label1.TabIndex = 1;
            this.label1.Text = "Anonymous usage logging is enabled by default in this version, this is very impor" +
    "tant for the development of this add-in.";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(142, 346);
            this.btnOk.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(112, 35);
            this.btnOk.TabIndex = 2;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 14);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(365, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Thank you for using the PowerPoint Warrior add-in!";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(14, 54);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(441, 54);
            this.label3.TabIndex = 4;
            this.label3.Text = "Brought to you by the Built to Lean team from Finland. Developer and contact pers" +
    "on: eric.selin@gmail.com";
            // 
            // tbEmail
            // 
            this.tbEmail.Location = new System.Drawing.Point(95, 117);
            this.tbEmail.Name = "tbEmail";
            this.tbEmail.Size = new System.Drawing.Size(316, 26);
            this.tbEmail.TabIndex = 5;
            this.tbEmail.WordWrap = false;
            this.tbEmail.Validating += new System.ComponentModel.CancelEventHandler(this.tbEmail_Validating);
            this.tbEmail.Validated += new System.EventHandler(this.tbEmail_Validated);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(18, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 20);
            this.label4.TabIndex = 6;
            this.label4.Text = "E-mail*:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(18, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(397, 40);
            this.label5.TabIndex = 7;
            this.label5.Text = "E-mail is required to use the PowerPoint Warrior toolbar\r\nPlease use the same e-m" +
    "ail that you signed up with";
            // 
            // errorpSettings
            // 
            this.errorpSettings.ContainerControl = this;
            // 
            // FormSettings
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 404);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbEmail);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.chkLogging);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormSettings";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings and About";
            ((System.ComponentModel.ISupportInitialize)(this.errorpSettings)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkLogging;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbEmail;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ErrorProvider errorpSettings;
    }
}