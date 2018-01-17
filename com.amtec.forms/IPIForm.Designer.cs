namespace ScreenPrinter.com.amtec.forms
{
    partial class IPIForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IPIForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbltime = new System.Windows.Forms.Label();
            this.lblLastCheck = new System.Windows.Forms.Label();
            this.lblIPITitle = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnAllowProduction = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnAnother = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LemonChiffon;
            this.panel1.Controls.Add(this.lbltime);
            this.panel1.Controls.Add(this.lblLastCheck);
            this.panel1.Controls.Add(this.lblIPITitle);
            this.panel1.Location = new System.Drawing.Point(-4, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(441, 106);
            this.panel1.TabIndex = 0;
            // 
            // lbltime
            // 
            this.lbltime.AutoSize = true;
            this.lbltime.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lbltime.Location = new System.Drawing.Point(152, 61);
            this.lbltime.Name = "lbltime";
            this.lbltime.Size = new System.Drawing.Size(128, 17);
            this.lbltime.TabIndex = 2;
            this.lbltime.Text = "2016-06-27 10:25:00";
            // 
            // lblLastCheck
            // 
            this.lblLastCheck.AutoSize = true;
            this.lblLastCheck.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblLastCheck.Location = new System.Drawing.Point(26, 61);
            this.lblLastCheck.Name = "lblLastCheck";
            this.lblLastCheck.Size = new System.Drawing.Size(77, 17);
            this.lblLastCheck.TabIndex = 1;
            this.lblLastCheck.Text = "Last Check:";
            // 
            // lblIPITitle
            // 
            this.lblIPITitle.AutoSize = true;
            this.lblIPITitle.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblIPITitle.Location = new System.Drawing.Point(26, 25);
            this.lblIPITitle.Name = "lblIPITitle";
            this.lblIPITitle.Size = new System.Drawing.Size(242, 17);
            this.lblIPITitle.TabIndex = 0;
            this.lblIPITitle.Text = "Initial Product Inspection in Process...";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnAllowProduction);
            this.panel2.Controls.Add(this.btnOk);
            this.panel2.Controls.Add(this.btnAnother);
            this.panel2.Location = new System.Drawing.Point(-1, 107);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(435, 61);
            this.panel2.TabIndex = 3;
            // 
            // btnAllowProduction
            // 
            this.btnAllowProduction.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAllowProduction.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.btnAllowProduction.Location = new System.Drawing.Point(129, 19);
            this.btnAllowProduction.Name = "btnAllowProduction";
            this.btnAllowProduction.Size = new System.Drawing.Size(136, 28);
            this.btnAllowProduction.TabIndex = 3;
            this.btnAllowProduction.Text = "AllowProduction";
            this.btnAllowProduction.UseVisualStyleBackColor = true;
            this.btnAllowProduction.Visible = false;
            this.btnAllowProduction.Click += new System.EventHandler(this.btnAllowProduction_Click);
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.btnOk.Location = new System.Drawing.Point(297, 19);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(136, 28);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Visible = false;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnAnother
            // 
            this.btnAnother.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAnother.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.btnAnother.Location = new System.Drawing.Point(155, 19);
            this.btnAnother.Name = "btnAnother";
            this.btnAnother.Size = new System.Drawing.Size(136, 28);
            this.btnAnother.TabIndex = 2;
            this.btnAnother.Text = "Another Board";
            this.btnAnother.UseVisualStyleBackColor = true;
            this.btnAnother.Visible = false;
            this.btnAnother.Click += new System.EventHandler(this.btnAnother_Click);
            // 
            // IPIForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(435, 169);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "IPIForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "IPI Status";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.IPIForm_FormClosing);
            this.Load += new System.EventHandler(this.IPIForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblLastCheck;
        public System.Windows.Forms.Label lblIPITitle;
        public System.Windows.Forms.Label lbltime;
        public System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.Button btnAllowProduction;
        public System.Windows.Forms.Button btnOk;
        public System.Windows.Forms.Button btnAnother;
    }
}