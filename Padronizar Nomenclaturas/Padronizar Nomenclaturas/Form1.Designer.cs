namespace Padronizar_Nomenclaturas
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSelecionePlanilha = new System.Windows.Forms.TextBox();
            this.buttonProcurar = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBoxDePara = new System.Windows.Forms.ComboBox();
            this.comboBoxColunaFabricante = new System.Windows.Forms.ComboBox();
            this.comboBoxColunaModeloPadrao = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnIniciar = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxSalvarOnde = new System.Windows.Forms.TextBox();
            this.buttonSelecionar = new System.Windows.Forms.Button();
            this.openFileDialogSelecionarPlanilha = new System.Windows.Forms.OpenFileDialog();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBoxColunaLeitura = new System.Windows.Forms.ComboBox();
            this.comboBoxColunaEscrita = new System.Windows.Forms.ComboBox();
            this.comboBoxAbaLeituraEscrita = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBoxColunaFab = new System.Windows.Forms.ComboBox();
            this.labelColunaFabricante = new System.Windows.Forms.Label();
            this.folderBrowserDialogSave = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selecione a planilha";
            // 
            // textBoxSelecionePlanilha
            // 
            this.textBoxSelecionePlanilha.Location = new System.Drawing.Point(16, 30);
            this.textBoxSelecionePlanilha.Name = "textBoxSelecionePlanilha";
            this.textBoxSelecionePlanilha.Size = new System.Drawing.Size(371, 20);
            this.textBoxSelecionePlanilha.TabIndex = 1;
            // 
            // buttonProcurar
            // 
            this.buttonProcurar.Location = new System.Drawing.Point(410, 28);
            this.buttonProcurar.Name = "buttonProcurar";
            this.buttonProcurar.Size = new System.Drawing.Size(90, 23);
            this.buttonProcurar.TabIndex = 2;
            this.buttonProcurar.Text = "Procurar";
            this.buttonProcurar.UseVisualStyleBackColor = true;
            this.buttonProcurar.Click += new System.EventHandler(this.buttonProcurar_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 185);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(165, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Selecione a aba com o \"DePara\"";
            // 
            // comboBoxDePara
            // 
            this.comboBoxDePara.FormattingEnabled = true;
            this.comboBoxDePara.Location = new System.Drawing.Point(16, 204);
            this.comboBoxDePara.Name = "comboBoxDePara";
            this.comboBoxDePara.Size = new System.Drawing.Size(481, 21);
            this.comboBoxDePara.TabIndex = 4;
            this.comboBoxDePara.SelectionChangeCommitted += new System.EventHandler(this.comboBoxDePara_SelectionChangeCommitted);
            // 
            // comboBoxColunaFabricante
            // 
            this.comboBoxColunaFabricante.FormattingEnabled = true;
            this.comboBoxColunaFabricante.Location = new System.Drawing.Point(16, 263);
            this.comboBoxColunaFabricante.Name = "comboBoxColunaFabricante";
            this.comboBoxColunaFabricante.Size = new System.Drawing.Size(230, 21);
            this.comboBoxColunaFabricante.TabIndex = 5;
            // 
            // comboBoxColunaModeloPadrao
            // 
            this.comboBoxColunaModeloPadrao.FormattingEnabled = true;
            this.comboBoxColunaModeloPadrao.Location = new System.Drawing.Point(273, 263);
            this.comboBoxColunaModeloPadrao.Name = "comboBoxColunaModeloPadrao";
            this.comboBoxColunaModeloPadrao.Size = new System.Drawing.Size(224, 21);
            this.comboBoxColunaModeloPadrao.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 243);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(128, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Coluna Fabricante/Marca";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(273, 244);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Coluna Modelo Padrão";
            // 
            // btnIniciar
            // 
            this.btnIniciar.Location = new System.Drawing.Point(172, 363);
            this.btnIniciar.Name = "btnIniciar";
            this.btnIniciar.Size = new System.Drawing.Size(165, 23);
            this.btnIniciar.TabIndex = 9;
            this.btnIniciar.Text = "Iniciar";
            this.btnIniciar.UseVisualStyleBackColor = true;
            this.btnIniciar.Click += new System.EventHandler(this.btnIniciar_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 299);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(143, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Selecione o local para salvar";
            // 
            // textBoxSalvarOnde
            // 
            this.textBoxSalvarOnde.Location = new System.Drawing.Point(16, 319);
            this.textBoxSalvarOnde.Name = "textBoxSalvarOnde";
            this.textBoxSalvarOnde.Size = new System.Drawing.Size(371, 20);
            this.textBoxSalvarOnde.TabIndex = 11;
            // 
            // buttonSelecionar
            // 
            this.buttonSelecionar.Location = new System.Drawing.Point(410, 319);
            this.buttonSelecionar.Name = "buttonSelecionar";
            this.buttonSelecionar.Size = new System.Drawing.Size(90, 23);
            this.buttonSelecionar.TabIndex = 12;
            this.buttonSelecionar.Text = "Selecionar";
            this.buttonSelecionar.UseVisualStyleBackColor = true;
            this.buttonSelecionar.Click += new System.EventHandler(this.buttonSelecionar_Click);
            // 
            // openFileDialogSelecionarPlanilha
            // 
            this.openFileDialogSelecionarPlanilha.FileName = "openFileDialog1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(13, 129);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(90, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "Coluna de Leitura";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(192, 129);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(90, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Coluna de Escrita";
            // 
            // comboBoxColunaLeitura
            // 
            this.comboBoxColunaLeitura.FormattingEnabled = true;
            this.comboBoxColunaLeitura.Location = new System.Drawing.Point(16, 145);
            this.comboBoxColunaLeitura.Name = "comboBoxColunaLeitura";
            this.comboBoxColunaLeitura.Size = new System.Drawing.Size(135, 21);
            this.comboBoxColunaLeitura.TabIndex = 15;
            // 
            // comboBoxColunaEscrita
            // 
            this.comboBoxColunaEscrita.FormattingEnabled = true;
            this.comboBoxColunaEscrita.Location = new System.Drawing.Point(193, 146);
            this.comboBoxColunaEscrita.Name = "comboBoxColunaEscrita";
            this.comboBoxColunaEscrita.Size = new System.Drawing.Size(130, 21);
            this.comboBoxColunaEscrita.TabIndex = 16;
            // 
            // comboBoxAbaLeituraEscrita
            // 
            this.comboBoxAbaLeituraEscrita.FormattingEnabled = true;
            this.comboBoxAbaLeituraEscrita.Location = new System.Drawing.Point(13, 85);
            this.comboBoxAbaLeituraEscrita.Name = "comboBoxAbaLeituraEscrita";
            this.comboBoxAbaLeituraEscrita.Size = new System.Drawing.Size(484, 21);
            this.comboBoxAbaLeituraEscrita.TabIndex = 17;
            this.comboBoxAbaLeituraEscrita.SelectionChangeCommitted += new System.EventHandler(this.comboBoxAbaLeituraEscrita_SelectionChangeCommitted);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(13, 66);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(173, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Selecione a aba de leitura e escrita";
            // 
            // comboBoxColunaFab
            // 
            this.comboBoxColunaFab.FormattingEnabled = true;
            this.comboBoxColunaFab.Location = new System.Drawing.Point(367, 145);
            this.comboBoxColunaFab.Name = "comboBoxColunaFab";
            this.comboBoxColunaFab.Size = new System.Drawing.Size(130, 21);
            this.comboBoxColunaFab.TabIndex = 19;
            // 
            // labelColunaFabricante
            // 
            this.labelColunaFabricante.AutoSize = true;
            this.labelColunaFabricante.Location = new System.Drawing.Point(367, 128);
            this.labelColunaFabricante.Name = "labelColunaFabricante";
            this.labelColunaFabricante.Size = new System.Drawing.Size(93, 13);
            this.labelColunaFabricante.TabIndex = 20;
            this.labelColunaFabricante.Text = "Coluna Fabricante";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(512, 409);
            this.Controls.Add(this.labelColunaFabricante);
            this.Controls.Add(this.comboBoxColunaFab);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.comboBoxAbaLeituraEscrita);
            this.Controls.Add(this.comboBoxColunaEscrita);
            this.Controls.Add(this.comboBoxColunaLeitura);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.buttonSelecionar);
            this.Controls.Add(this.textBoxSalvarOnde);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnIniciar);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxColunaModeloPadrao);
            this.Controls.Add(this.comboBoxColunaFabricante);
            this.Controls.Add(this.comboBoxDePara);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonProcurar);
            this.Controls.Add(this.textBoxSelecionePlanilha);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Padronizador de Nomenclaturas";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxSelecionePlanilha;
        private System.Windows.Forms.Button buttonProcurar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxDePara;
        private System.Windows.Forms.ComboBox comboBoxColunaFabricante;
        private System.Windows.Forms.ComboBox comboBoxColunaModeloPadrao;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnIniciar;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxSalvarOnde;
        private System.Windows.Forms.Button buttonSelecionar;
        private System.Windows.Forms.OpenFileDialog openFileDialogSelecionarPlanilha;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBoxColunaLeitura;
        private System.Windows.Forms.ComboBox comboBoxColunaEscrita;
        private System.Windows.Forms.ComboBox comboBoxAbaLeituraEscrita;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBoxColunaFab;
        private System.Windows.Forms.Label labelColunaFabricante;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogSave;
    }
}

