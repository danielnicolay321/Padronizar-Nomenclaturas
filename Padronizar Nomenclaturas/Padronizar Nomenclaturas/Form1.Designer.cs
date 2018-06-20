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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.buttonSelecionar = new System.Windows.Forms.Button();
            this.openFileDialogSelecionarPlanilha = new System.Windows.Forms.OpenFileDialog();
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
            this.label2.Location = new System.Drawing.Point(13, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(165, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Selecione a aba com o \"DePara\"";
            // 
            // comboBoxDePara
            // 
            this.comboBoxDePara.FormattingEnabled = true;
            this.comboBoxDePara.Location = new System.Drawing.Point(16, 85);
            this.comboBoxDePara.Name = "comboBoxDePara";
            this.comboBoxDePara.Size = new System.Drawing.Size(481, 21);
            this.comboBoxDePara.TabIndex = 4;
            this.comboBoxDePara.SelectionChangeCommitted += new System.EventHandler(this.comboBoxDePara_SelectionChangeCommitted);
            // 
            // comboBoxColunaFabricante
            // 
            this.comboBoxColunaFabricante.FormattingEnabled = true;
            this.comboBoxColunaFabricante.Location = new System.Drawing.Point(16, 144);
            this.comboBoxColunaFabricante.Name = "comboBoxColunaFabricante";
            this.comboBoxColunaFabricante.Size = new System.Drawing.Size(208, 21);
            this.comboBoxColunaFabricante.TabIndex = 5;
            // 
            // comboBoxColunaModeloPadrao
            // 
            this.comboBoxColunaModeloPadrao.FormattingEnabled = true;
            this.comboBoxColunaModeloPadrao.Location = new System.Drawing.Point(273, 144);
            this.comboBoxColunaModeloPadrao.Name = "comboBoxColunaModeloPadrao";
            this.comboBoxColunaModeloPadrao.Size = new System.Drawing.Size(203, 21);
            this.comboBoxColunaModeloPadrao.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 124);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Coluna Fabricante";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(273, 125);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Coluna Modelo Padrão";
            // 
            // btnIniciar
            // 
            this.btnIniciar.Location = new System.Drawing.Point(171, 253);
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
            this.label5.Location = new System.Drawing.Point(16, 180);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(143, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Selecione o local para salvar";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(16, 200);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(371, 20);
            this.textBox1.TabIndex = 11;
            // 
            // buttonSelecionar
            // 
            this.buttonSelecionar.Location = new System.Drawing.Point(410, 200);
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
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(512, 296);
            this.Controls.Add(this.buttonSelecionar);
            this.Controls.Add(this.textBox1);
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
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button buttonSelecionar;
        private System.Windows.Forms.OpenFileDialog openFileDialogSelecionarPlanilha;
    }
}

