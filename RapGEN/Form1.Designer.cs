namespace RapGEN
{
    partial class RapGEN_MainWin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RapGEN_MainWin));
            this.InputPathLabel = new System.Windows.Forms.Label();
            this.OutputPathLabel = new System.Windows.Forms.Label();
            this.InputPathTB = new System.Windows.Forms.TextBox();
            this.BrowseButton1 = new System.Windows.Forms.Button();
            this.BrowseButton2 = new System.Windows.Forms.Button();
            this.OutputPathTB = new System.Windows.Forms.TextBox();
            this.GenButton = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progBarLabel1 = new System.Windows.Forms.Label();
            this.Credits = new System.Windows.Forms.Label();
            this.mainTimer = new System.Windows.Forms.Timer(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // InputPathLabel
            // 
            resources.ApplyResources(this.InputPathLabel, "InputPathLabel");
            this.InputPathLabel.BackColor = System.Drawing.Color.Transparent;
            this.InputPathLabel.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.InputPathLabel.Name = "InputPathLabel";
            // 
            // OutputPathLabel
            // 
            resources.ApplyResources(this.OutputPathLabel, "OutputPathLabel");
            this.OutputPathLabel.BackColor = System.Drawing.Color.Transparent;
            this.OutputPathLabel.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.OutputPathLabel.Name = "OutputPathLabel";
            // 
            // InputPathTB
            // 
            resources.ApplyResources(this.InputPathTB, "InputPathTB");
            this.InputPathTB.Name = "InputPathTB";
            this.InputPathTB.TextChanged += new System.EventHandler(this.InputPathTB_Changed);
            // 
            // BrowseButton1
            // 
            resources.ApplyResources(this.BrowseButton1, "BrowseButton1");
            this.BrowseButton1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.BrowseButton1.Name = "BrowseButton1";
            this.BrowseButton1.UseVisualStyleBackColor = false;
            this.BrowseButton1.Click += new System.EventHandler(this.BrowseButton1_Click);
            // 
            // BrowseButton2
            // 
            resources.ApplyResources(this.BrowseButton2, "BrowseButton2");
            this.BrowseButton2.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.BrowseButton2.Name = "BrowseButton2";
            this.BrowseButton2.UseVisualStyleBackColor = false;
            this.BrowseButton2.Click += new System.EventHandler(this.BrowseButton2_Click);
            // 
            // OutputPathTB
            // 
            resources.ApplyResources(this.OutputPathTB, "OutputPathTB");
            this.OutputPathTB.Name = "OutputPathTB";
            // 
            // GenButton
            // 
            resources.ApplyResources(this.GenButton, "GenButton");
            this.GenButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.GenButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GenButton.ForeColor = System.Drawing.Color.DarkGreen;
            this.GenButton.Name = "GenButton";
            this.GenButton.UseVisualStyleBackColor = false;
            this.GenButton.Click += new System.EventHandler(this.GenButton_Click);
            // 
            // progressBar1
            // 
            resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Cursor = System.Windows.Forms.Cursors.No;
            this.progressBar1.Name = "progressBar1";
            // 
            // progBarLabel1
            // 
            resources.ApplyResources(this.progBarLabel1, "progBarLabel1");
            this.progBarLabel1.BackColor = System.Drawing.Color.Transparent;
            this.progBarLabel1.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.progBarLabel1.Name = "progBarLabel1";
            // 
            // Credits
            // 
            resources.ApplyResources(this.Credits, "Credits");
            this.Credits.BackColor = System.Drawing.Color.Transparent;
            this.Credits.ForeColor = System.Drawing.Color.Gray;
            this.Credits.Name = "Credits";
            // 
            // mainTimer
            // 
            this.mainTimer.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "*.xlsx";
            resources.ApplyResources(this.saveFileDialog1, "saveFileDialog1");
            // 
            // RapGEN_MainWin
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.Credits);
            this.Controls.Add(this.progBarLabel1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.GenButton);
            this.Controls.Add(this.BrowseButton2);
            this.Controls.Add(this.OutputPathTB);
            this.Controls.Add(this.BrowseButton1);
            this.Controls.Add(this.InputPathTB);
            this.Controls.Add(this.OutputPathLabel);
            this.Controls.Add(this.InputPathLabel);
            this.Name = "RapGEN_MainWin";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label InputPathLabel;
        private System.Windows.Forms.Label OutputPathLabel;
        private System.Windows.Forms.TextBox InputPathTB;
        private System.Windows.Forms.Button BrowseButton1;
        private System.Windows.Forms.Button BrowseButton2;
        private System.Windows.Forms.TextBox OutputPathTB;
        private System.Windows.Forms.Button GenButton;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label progBarLabel1;
        private System.Windows.Forms.Label Credits;
        private System.Windows.Forms.Timer mainTimer;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

