namespace TestAutoGenerator
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
            this.btnProcess = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.cbApplications = new System.Windows.Forms.ComboBox();
            this.txtFile1 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnSelectFile_Homologation = new System.Windows.Forms.Button();
            this.txtHomologation = new System.Windows.Forms.TextBox();
            this.chKHomologation = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdGasoline = new System.Windows.Forms.RadioButton();
            this.rdDiesel = new System.Windows.Forms.RadioButton();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnConfig = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txtProgress = new System.Windows.Forms.RichTextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblProgress = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnProcess
            // 
            this.btnProcess.BackColor = System.Drawing.Color.DarkGray;
            this.btnProcess.ForeColor = System.Drawing.Color.Navy;
            this.btnProcess.Location = new System.Drawing.Point(11, 325);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(155, 32);
            this.btnProcess.TabIndex = 0;
            this.btnProcess.Text = "Generate template";
            this.btnProcess.UseVisualStyleBackColor = false;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.Controls.Add(this.groupBox3);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.btnNew);
            this.panel1.Controls.Add(this.btnConfig);
            this.panel1.Controls.Add(this.btnProcess);
            this.panel1.Location = new System.Drawing.Point(38, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(769, 375);
            this.panel1.TabIndex = 1;
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.SystemColors.ControlLight;
            this.groupBox3.Controls.Add(this.btnSelectFile);
            this.groupBox3.Controls.Add(this.cbApplications);
            this.groupBox3.Controls.Add(this.txtFile1);
            this.groupBox3.Location = new System.Drawing.Point(11, 10);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(747, 107);
            this.groupBox3.TabIndex = 17;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "SdF";
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.BackColor = System.Drawing.Color.DarkGray;
            this.btnSelectFile.ForeColor = System.Drawing.Color.Navy;
            this.btnSelectFile.Location = new System.Drawing.Point(17, 21);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(138, 28);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "Select file";
            this.btnSelectFile.UseVisualStyleBackColor = false;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // cbApplications
            // 
            this.cbApplications.FormattingEnabled = true;
            this.cbApplications.Location = new System.Drawing.Point(188, 63);
            this.cbApplications.Name = "cbApplications";
            this.cbApplications.Size = new System.Drawing.Size(534, 24);
            this.cbApplications.TabIndex = 3;
            this.cbApplications.Text = "Select application";
            // 
            // txtFile1
            // 
            this.txtFile1.Enabled = false;
            this.txtFile1.Location = new System.Drawing.Point(188, 21);
            this.txtFile1.Margin = new System.Windows.Forms.Padding(0);
            this.txtFile1.MinimumSize = new System.Drawing.Size(4, 28);
            this.txtFile1.Name = "txtFile1";
            this.txtFile1.Size = new System.Drawing.Size(534, 22);
            this.txtFile1.TabIndex = 2;
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.ControlLight;
            this.groupBox2.Controls.Add(this.btnSelectFile_Homologation);
            this.groupBox2.Controls.Add(this.txtHomologation);
            this.groupBox2.Controls.Add(this.chKHomologation);
            this.groupBox2.Location = new System.Drawing.Point(11, 126);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(747, 114);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "TdH";
            // 
            // btnSelectFile_Homologation
            // 
            this.btnSelectFile_Homologation.BackColor = System.Drawing.Color.DarkGray;
            this.btnSelectFile_Homologation.Enabled = false;
            this.btnSelectFile_Homologation.ForeColor = System.Drawing.Color.Navy;
            this.btnSelectFile_Homologation.Location = new System.Drawing.Point(17, 68);
            this.btnSelectFile_Homologation.Name = "btnSelectFile_Homologation";
            this.btnSelectFile_Homologation.Size = new System.Drawing.Size(139, 28);
            this.btnSelectFile_Homologation.TabIndex = 11;
            this.btnSelectFile_Homologation.Text = "Select file";
            this.btnSelectFile_Homologation.UseVisualStyleBackColor = false;
            this.btnSelectFile_Homologation.Click += new System.EventHandler(this.btnSelectFile_Homologation_Click);
            // 
            // txtHomologation
            // 
            this.txtHomologation.Enabled = false;
            this.txtHomologation.Location = new System.Drawing.Point(188, 68);
            this.txtHomologation.Margin = new System.Windows.Forms.Padding(0);
            this.txtHomologation.MinimumSize = new System.Drawing.Size(4, 28);
            this.txtHomologation.Name = "txtHomologation";
            this.txtHomologation.Size = new System.Drawing.Size(534, 22);
            this.txtHomologation.TabIndex = 12;
            // 
            // chKHomologation
            // 
            this.chKHomologation.AutoSize = true;
            this.chKHomologation.Location = new System.Drawing.Point(17, 30);
            this.chKHomologation.Name = "chKHomologation";
            this.chKHomologation.Size = new System.Drawing.Size(151, 21);
            this.chKHomologation.TabIndex = 10;
            this.chKHomologation.Text = "MIL Check Enabled";
            this.chKHomologation.UseVisualStyleBackColor = true;
            this.chKHomologation.CheckedChanged += new System.EventHandler(this.chKHomologation_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.groupBox1.Controls.Add(this.rdGasoline);
            this.groupBox1.Controls.Add(this.rdDiesel);
            this.groupBox1.Location = new System.Drawing.Point(11, 250);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(343, 68);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Fuel type";
            // 
            // rdGasoline
            // 
            this.rdGasoline.AutoSize = true;
            this.rdGasoline.Location = new System.Drawing.Point(17, 30);
            this.rdGasoline.Name = "rdGasoline";
            this.rdGasoline.Size = new System.Drawing.Size(85, 21);
            this.rdGasoline.TabIndex = 13;
            this.rdGasoline.TabStop = true;
            this.rdGasoline.Text = "Gasoline";
            this.rdGasoline.UseVisualStyleBackColor = true;
            // 
            // rdDiesel
            // 
            this.rdDiesel.AutoSize = true;
            this.rdDiesel.Location = new System.Drawing.Point(188, 30);
            this.rdDiesel.Name = "rdDiesel";
            this.rdDiesel.Size = new System.Drawing.Size(68, 21);
            this.rdDiesel.TabIndex = 14;
            this.rdDiesel.TabStop = true;
            this.rdDiesel.Text = "Diesel";
            this.rdDiesel.UseVisualStyleBackColor = true;
            // 
            // btnNew
            // 
            this.btnNew.BackColor = System.Drawing.Color.DarkGray;
            this.btnNew.Enabled = false;
            this.btnNew.ForeColor = System.Drawing.Color.Navy;
            this.btnNew.Location = new System.Drawing.Point(199, 324);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(155, 32);
            this.btnNew.TabIndex = 9;
            this.btnNew.Text = "New run";
            this.btnNew.UseVisualStyleBackColor = false;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnConfig
            // 
            this.btnConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnConfig.BackColor = System.Drawing.Color.DarkGray;
            this.btnConfig.ForeColor = System.Drawing.Color.Navy;
            this.btnConfig.Location = new System.Drawing.Point(603, 324);
            this.btnConfig.Name = "btnConfig";
            this.btnConfig.Size = new System.Drawing.Size(155, 32);
            this.btnConfig.TabIndex = 4;
            this.btnConfig.Text = "App Config";
            this.btnConfig.UseVisualStyleBackColor = false;
            this.btnConfig.Visible = false;
            this.btnConfig.Click += new System.EventHandler(this.btnConfig_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // panel2
            // 
            this.panel2.AutoSize = true;
            this.panel2.BackColor = System.Drawing.SystemColors.Control;
            this.panel2.Controls.Add(this.txtProgress);
            this.panel2.Controls.Add(this.btnCancel);
            this.panel2.Controls.Add(this.lblProgress);
            this.panel2.Controls.Add(this.progressBar1);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(38, 405);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(769, 389);
            this.panel2.TabIndex = 2;
            // 
            // txtProgress
            // 
            this.txtProgress.BackColor = System.Drawing.Color.White;
            this.txtProgress.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtProgress.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProgress.ForeColor = System.Drawing.Color.Black;
            this.txtProgress.HideSelection = false;
            this.txtProgress.Location = new System.Drawing.Point(12, 84);
            this.txtProgress.Name = "txtProgress";
            this.txtProgress.Size = new System.Drawing.Size(745, 292);
            this.txtProgress.TabIndex = 13;
            this.txtProgress.Text = "";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.DarkGray;
            this.btnCancel.ForeColor = System.Drawing.Color.Navy;
            this.btnCancel.Location = new System.Drawing.Point(602, 17);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(155, 32);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.BackColor = System.Drawing.Color.DarkGray;
            this.lblProgress.ForeColor = System.Drawing.Color.Navy;
            this.lblProgress.Location = new System.Drawing.Point(15, 55);
            this.lblProgress.MinimumSize = new System.Drawing.Size(38, 0);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Padding = new System.Windows.Forms.Padding(3);
            this.lblProgress.Size = new System.Drawing.Size(38, 23);
            this.lblProgress.TabIndex = 12;
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.Black;
            this.progressBar1.ForeColor = System.Drawing.Color.DarkGray;
            this.progressBar1.Location = new System.Drawing.Point(59, 55);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(698, 23);
            this.progressBar1.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.DarkGray;
            this.label1.ForeColor = System.Drawing.Color.Navy;
            this.label1.Location = new System.Drawing.Point(15, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label1.Name = "label1";
            this.label1.Padding = new System.Windows.Forms.Padding(7);
            this.label1.Size = new System.Drawing.Size(134, 31);
            this.label1.TabIndex = 10;
            this.label1.Text = "Processing status";
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(49, 399);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(746, 2);
            this.label2.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 811);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.Text = "Diag automatic test generator";
            this.panel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtFile1;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.ComboBox cbApplications;
        private System.Windows.Forms.Button btnConfig;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.RichTextBox txtProgress;
        private System.Windows.Forms.TextBox txtHomologation;
        private System.Windows.Forms.Button btnSelectFile_Homologation;
        private System.Windows.Forms.CheckBox chKHomologation;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdGasoline;
        private System.Windows.Forms.RadioButton rdDiesel;
        private System.Windows.Forms.Label label2;
    }
}

