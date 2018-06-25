using Padronizar_Nomenclaturas.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Padronizar_Nomenclaturas
{
    public partial class Form1 : Form
    {
        DoThings doThings;
        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void buttonProcurar_Click(object sender, EventArgs e)
        {
            btnIniciar.Enabled = false;
            textBoxSelecionePlanilha.Text = ChooseFolder(openFileDialogSelecionarPlanilha);

            try
            {
                doThings = new DoThings(textBoxSelecionePlanilha.Text);

                comboBoxDePara.DataSource = new BindingSource(doThings.Abas, null);
                comboBoxDePara.SelectedIndex = -1;
                comboBoxDePara.DisplayMember = "Key";
                comboBoxDePara.ValueMember = "Value";

                comboBoxAbaLeituraEscrita.DataSource = new BindingSource(doThings.Abas, null);
                comboBoxAbaLeituraEscrita.SelectedIndex = -1;
                comboBoxAbaLeituraEscrita.DisplayMember = "Key";
                comboBoxAbaLeituraEscrita.ValueMember = "Value";
            }
            catch(Exception ex)
            {
                Console.WriteLine("");
            }
        }

        // selecionar local para salvar
        private void buttonSelecionar_Click(object sender, EventArgs e)
        {
            textBoxSalvarOnde.Text = SaveFolder(folderBrowserDialogSave);
            
        }

        private void comboBoxAbaLeituraEscrita_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                var result = doThings.ValidarEntradasAbaDePara(comboBoxAbaLeituraEscrita.GetItemText(comboBoxDePara.SelectedItem));

                //  var colunas = result.Where(x => x.Key.Equals(comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem))).Select(z => z.Value);
                var colunas = result.First(x => x.Key == comboBoxAbaLeituraEscrita.GetItemText(comboBoxDePara.SelectedItem)).Value;//.ToDictionary(dc => dc.Key, dc => dc.Value);

                comboBoxColunaLeitura.DataSource = new BindingSource(colunas, null);
                comboBoxColunaLeitura.DisplayMember = "Value";
                comboBoxColunaLeitura.ValueMember = "Key";

                comboBoxColunaEscrita.DataSource = new BindingSource(colunas, null);
                comboBoxColunaEscrita.DisplayMember = "Value";
                comboBoxColunaEscrita.ValueMember = "Key";

                comboBoxColunaFab.DataSource = new BindingSource(colunas, null);
                comboBoxColunaFab.DisplayMember = "Value";
                comboBoxColunaFab.ValueMember = "Key";

                btnIniciar.Enabled = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
            }

        }


        private void comboBoxDePara_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                var result = doThings.ValidarEntradasAbaDePara(comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem));

              //  var colunas = result.Where(x => x.Key.Equals(comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem))).Select(z => z.Value);
                var colunas = result.First(x => x.Key == comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem)).Value;//.ToDictionary(dc => dc.Key, dc => dc.Value);

                comboBoxColunaFabricante.DataSource = new BindingSource(colunas, null);
                comboBoxColunaFabricante.DisplayMember = "Value";
                comboBoxColunaFabricante.ValueMember = "Key";

                comboBoxColunaModeloPadrao.DataSource = new BindingSource(colunas, null);
                comboBoxColunaModeloPadrao.DisplayMember = "Value";
                comboBoxColunaModeloPadrao.ValueMember = "Key";

                btnIniciar.Enabled = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
            }
        }

        public string ChooseFolder(OpenFileDialog openfiledialog)
        {
            //lblStatus.Text = "escolhendo arquivo...";
            if (openfiledialog.ShowDialog() == DialogResult.OK)
            {
                return openfiledialog.FileName;
            }
            return "";
        }

        public string SaveFolder(FolderBrowserDialog folderBrowserDialogSave)
        {
            if (folderBrowserDialogSave.ShowDialog() == DialogResult.OK)
            {
                return folderBrowserDialogSave.SelectedPath;
            }
            return "";
        }

       

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            var values = doThings.CarregarValoresDePara(comboBoxColunaFabricante.GetItemText(comboBoxColunaFabricante.SelectedItem),
              comboBoxColunaModeloPadrao.GetItemText(comboBoxColunaModeloPadrao.SelectedItem),
                comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem));

            var sheetName = comboBoxAbaLeituraEscrita.GetItemText(comboBoxAbaLeituraEscrita.SelectedItem);
            var columnToRead = comboBoxColunaLeitura.GetItemText(comboBoxColunaLeitura.SelectedItem);
            var columnToWrite = comboBoxColunaEscrita.GetItemText(comboBoxColunaEscrita.SelectedItem);
            var columnFabricante = comboBoxColunaFab.GetItemText(comboBoxColunaFab.SelectedItem);
            var path = textBoxSalvarOnde.Text;
            //string sheetName, string columnToRead, string columnToWrite

            doThings.SetColumnIndexOfReadAndWrite(sheetName, columnToRead, columnToWrite, columnFabricante);
            doThings.WriteExcelFile(sheetName, columnToRead, columnToWrite, path);

        }

      



  


    }
}
