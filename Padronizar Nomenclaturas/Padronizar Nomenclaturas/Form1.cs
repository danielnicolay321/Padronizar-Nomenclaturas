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
            textBoxSelecionePlanilha.Text = ChooseFolder();

            try
            {
                doThings = new DoThings(textBoxSelecionePlanilha.Text);

                comboBoxDePara.DataSource = new BindingSource(doThings.Abas, null);
                comboBoxDePara.DisplayMember = "Key";
                comboBoxDePara.ValueMember = "Value";            
            }
            catch(Exception ex)
            {
                Console.WriteLine("");
            }
        }

        private void comboBoxDePara_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                var colunas = doThings.ValidarEntradasAbaDePara(comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem));

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

        public string ChooseFolder()
        {
            //lblStatus.Text = "escolhendo arquivo...";
            if (openFileDialogSelecionarPlanilha.ShowDialog() == DialogResult.OK)
            {
                return openFileDialogSelecionarPlanilha.FileName;
            }
            return "";
        }

        // selecionar local para salvar
        private void buttonSelecionar_Click(object sender, EventArgs e)
        {

        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            var values = doThings.CarregarValoresDePara(comboBoxColunaFabricante.GetItemText(comboBoxColunaFabricante.SelectedItem),
              comboBoxColunaModeloPadrao.GetItemText(comboBoxColunaModeloPadrao.SelectedItem),
                comboBoxDePara.GetItemText(comboBoxDePara.SelectedItem));
        }



        #region WriteExcelFile        
        //public static void WriteExcelFile(System.Data.DataTable dataTable)
        //{
        //    Microsoft.Office.Interop.Excel.Application excel;
        //    Microsoft.Office.Interop.Excel.Workbook worKbooK;
        //    Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
        //    Microsoft.Office.Interop.Excel.Range celLrangE;

        //    excel = new Microsoft.Office.Interop.Excel.Application();
        //    excel.Visible = false;
        //    excel.DisplayAlerts = false;
        //    worKbooK = excel.Workbooks.Add(Type.Missing);
        //    worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
        //    int rowcount = 2;

        //    foreach (DataRow datarow in dataTable.Rows)
        //    {
        //        rowcount += 1;
        //        for (int i = 1; i <= dataTable.Columns.Count; i++)
        //        {

        //            if (rowcount == 3)
        //            {
        //                worKsheeT.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;

        //            }

        //            worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

        //            if (rowcount > 3)
        //            {
        //                if (i == dataTable.Columns.Count)
        //                {
        //                    if (rowcount % 2 == 0)
        //                    {
        //                        celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, dataTable.Columns.Count]];

        //                    }
        //                }
        //            } // if (rowcount > 3)
        //        } // for (int i = 1; i <= tst.Columns.Count; i++)
        //    } // foreach (DataRow datarow in tst.Rows)


        //    celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, dataTable.Columns.Count]];
        //    celLrangE.EntireColumn.AutoFit();
        //    celLrangE.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        //    celLrangE.Font.Size = 8;
        //    celLrangE.Font.FontStyle = "Calibri";

        //    Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
        //    border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        //    border.Weight = 2d;

        //    celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, dataTable.Columns.Count]];

        //    var path = @"C:\Relatorio";

        //    try
        //    {
        //        // Determine whether the directory exists.
        //        if (!Directory.Exists(path))
        //        {
        //            // Try to create the directory.
        //            DirectoryInfo di = Directory.CreateDirectory(path);
        //        }

        //        worKbooK.SaveAs(path + "\\planilha_semanal_Robo.xlsx");
        //        worKbooK.Close();

        //        // Delete the directory.
        //        //di.Delete();
        //        //Console.WriteLine("The directory was deleted successfully.");
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("The process to check/create folder failed: {0}", e.ToString());
        //    }

        //    excel.Quit();


        //    Console.WriteLine("Planilha gerada com sucesso. Ela foi salva no diretorio -> C:/Relatorio/planilha_semanal_Robo.xlsx");

        //    Console.WriteLine("Aperte a tecla ENTER para finalizar o programa....");
        //    ConsoleKeyInfo keyInfo = Console.ReadKey();

        //    while (keyInfo.Key != ConsoleKey.Enter)
        //        keyInfo = Console.ReadKey();
        //}
        #endregion


    }
}
