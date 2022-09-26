namespace AutomisationHospitalData
{
    partial class Form1
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
            this.hørkramButton = new System.Windows.Forms.Button();
            this.hørkramTextbox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // hørkramButton
            // 
            this.hørkramButton.Location = new System.Drawing.Point(205, 84);
            this.hørkramButton.Name = "hørkramButton";
            this.hørkramButton.Size = new System.Drawing.Size(75, 23);
            this.hørkramButton.TabIndex = 0;
            this.hørkramButton.Text = "Hørkram";
            this.hørkramButton.UseVisualStyleBackColor = true;
            this.hørkramButton.Click += new System.EventHandler(this.hørkramButton_Click);
            // 
            // hørkramTextbox
            // 
            this.hørkramTextbox.Location = new System.Drawing.Point(99, 86);
            this.hørkramTextbox.Name = "hørkramTextbox";
            this.hørkramTextbox.Size = new System.Drawing.Size(100, 20);
            this.hørkramTextbox.TabIndex = 1;
            this.hørkramTextbox.TextChanged += new System.EventHandler(this.hørkramTextbox_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.hørkramTextbox);
            this.Controls.Add(this.hørkramButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button hørkramButton;
        private System.Windows.Forms.TextBox hørkramTextbox;
    }
}

