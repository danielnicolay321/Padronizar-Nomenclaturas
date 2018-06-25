using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Padronizar_Nomenclaturas.Classes
{
    public class DoThings
    {
        private static string _planilhaDeOrigem;
        private static string _planilhaParaSalvar;

        //public int ColFabricante { get; set; }
        //public int ColModeloPadrao { get; set; }

        public Excel._Application oApp { get; set; }
        public Excel.Workbook oWorkbook { get; set; }
        public Excel.Workbook oWorkbookCopy { get; set; }
        public Excel.Worksheet Wsheet { get; set; }
        public List<KeyValuePair<string, Dictionary<int, string>>> DeParaColumns { get; set; }
      //  public Dictionary<int, string> DeParaColumns { get; set; }
        public Dictionary<int, string> ReadAndWriteColumns { get; set; }
        public Dictionary<string, int> Abas { get; set; }
        public List<KeyValuePair<string, string>> ValoresDePara { get; set; }
        public List<KeyValuePair<string, List<string>>> wordListByBrand { get; set; }
        public List<KeyValuePair<string, List<string>>> modelsByBrand { get; set; }
        public Sheets sheets { get; set; }

        Tuple<int, int> IndiceColunaModeloPadrao = new Tuple<int, int>(0, 0);
        Tuple<int, int> IndiceColunaFabricante = new Tuple<int, int>(0, 0);
        Tuple<int, int> IndiceColunaFabricante2 = new Tuple<int, int>(0, 0);
        Tuple<int, int> IndiceColunaParaLer = new Tuple<int, int>(0, 0);
        Tuple<int, int> IndiceColunaParaGravar = new Tuple<int, int>(0, 0);

        public DoThings(string nomeItem)
        {
            DeParaColumns = new List<KeyValuePair<string, Dictionary<int, string>>>();
            _planilhaDeOrigem = nomeItem;
            ValidarEntradasPlanilhaOrigem();      
        }

        public void ValidarEntradasPlanilhaOrigem()
        {
            //verifica se o caminho da planilha de origem é valido
            if (string.IsNullOrEmpty(_planilhaDeOrigem))
            {   //se nao for válido, dispara exception
                throw new ArgumentException("Selecione uma planilha válida");
            }
            // instancia uma nova instancia de excel application
            oApp = new Excel.Application();
            //verifica se a aplicação excel foi instanciada
            if (oApp == null)
            {   //se nao for, vai disparar exception
                throw new ApplicationException("Excel não pôde ser iniciado. Verifique se ele está devidamente instalado em seu computador");
            }

            oApp.Visible = false;
            oWorkbook = oApp.Workbooks.Open(_planilhaDeOrigem);

            sheets = oWorkbook.Worksheets;
            Abas = new Dictionary<string, int>();
            var indice = 1;

            foreach(Worksheet worksheet in oWorkbook.Worksheets)
            {
                Abas.Add(worksheet.Name, indice);
                indice++;
            }
            indice = 1;
        }

        public List<KeyValuePair<string, Dictionary<int, string>>> ValidarEntradasAbaDePara(string nomeAba)
        {
           // new KeyValuePair<string, Dictionary<int, List<string>>>();
            var columnsDictionary = new Dictionary<int, List<string>>();

            var columns = new Dictionary<int,string>();
            var table = oWorkbook.Worksheets[nomeAba];
            var colNo = table.UsedRange.Columns.Count;
            object[,] array = table.UsedRange.Value;
           
            for (var indice = 1; indice <= colNo; indice++)
            {
                if (array[1, indice] != null)
                {
                    columns.Add(indice, array[1, indice].ToString());
                }
            }

            DeParaColumns.Add(new KeyValuePair<string, Dictionary<int, string>>(nomeAba, columns));

            return DeParaColumns;
        }

        public List<KeyValuePair<string, string>> CarregarValoresDePara(string ColunaFabricante, string ColunaModeloPadrao, string nomeAba)
        {
            ValoresDePara = new List<KeyValuePair<string, string>>();
            var excelSheet = oWorkbook.Worksheets[nomeAba];
            var colNo = excelSheet.UsedRange.Columns.Count;
            var rowNo = excelSheet.UsedRange.Rows.Count;
            object[,] array = excelSheet.UsedRange.Value;

            for (int i = 1; i <= colNo; i++ )
            {
                if (array[1, i] != null)
                {
                    if (array[1, i].ToString() == ColunaFabricante)
                    {
                        IndiceColunaFabricante = new Tuple<int, int>(1, i);
                    }
                    else if(array[1, i].ToString() == ColunaModeloPadrao)
                    {
                        IndiceColunaModeloPadrao = new Tuple<int, int>(1, i);                       
                    }
                }
            } // for  (int i = 1; i <= colNo; i++ )

            for (int i = 1; i <= rowNo; i++)
            {
                int contAux = i + 1;
                if (contAux > rowNo)
                {
                    break;
                }
                try
                {
                    ValoresDePara.Add(new KeyValuePair<string, string>((array[contAux, IndiceColunaFabricante.Item2].ToString()), (array[contAux, IndiceColunaModeloPadrao.Item2].ToString())));
                }
                catch (Exception e)
                {

                }
            }

            var valoresDeParaAgrupadosPorFabricante = ValoresDePara.GroupBy(x => x.Key)
                .Select(group => new KeyValuePair<string, List<string>>(group.Key, group.Select(t => t.Value).ToList())).ToList();
            SepararModelosPorFabricante(valoresDeParaAgrupadosPorFabricante);

            return ValoresDePara;
        } // public Dictionary<string, string> CarregarValoresDePara

        public void SetColumnIndexOfReadAndWrite(string sheetSname, string columnToRead, string columnToWrite, string columnFabricante)
        {
            var excelSheet = oWorkbook.Worksheets[sheetSname];
            var colNo = excelSheet.UsedRange.Columns.Count;
            var rowNo = excelSheet.UsedRange.Rows.Count;
            object[,] array = excelSheet.UsedRange.Value;

            for (int i = 1; i <= colNo; i++)
            {
                if (array[1, i] != null)
                {
                    if (array[1, i].ToString() == columnToRead)
                    {
                        IndiceColunaParaLer = new Tuple<int, int>(1, i);
                    }
                    else if (array[1, i].ToString() == columnToWrite)
                    {
                        IndiceColunaParaGravar = new Tuple<int, int>(1, i);
                    }
                    else if (array[1, i].ToString() == columnFabricante)
                    {
                        IndiceColunaFabricante2 = new Tuple<int, int>(1, i);
                    }
                }
            } // for  (int i = 1; i <= colNo; i++ )
        }

        public void SepararModelosPorFabricante(List<KeyValuePair<string, List<string>>> modelsByBrand1)
        {
            modelsByBrand = modelsByBrand1;
            wordListByBrand = new List<KeyValuePair<string, List<string>>>();

            foreach(var item in modelsByBrand1)
            {
                var itemList = new List<string>();
                foreach(var itemValue in item.Value)
                {
                    string[] splittedWords = itemValue.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    itemList.AddRange(splittedWords);
                }

                var itemsOccurrences = itemList.GroupBy(c => c)
                    .ToDictionary(grp => grp.Key, grp => grp.Count());

                var itemListWithoutRepetition = itemsOccurrences.Keys.ToList();

                wordListByBrand.Add(new KeyValuePair<string, List<string>>(item.Key, itemListWithoutRepetition));

            } // foreach(var item in modelsByBrand)
            Console.WriteLine("");
            //GetModelNameByBrand("iphone 5s 16gb adas adoijsa motorola samsung teste ifone 4g", "APPLE");
        } // public void SepararModelosPorFabricante(List<KeyValuePair<string, List<string>>> modelsByBrand)


        public string GetModelNameByBrand(string text, string brand)
        {
            var arrImportante = wordListByBrand.Where(x => x.Key.ToLower().Equals(brand.ToLower())).Select(y => y.Value);

            List<string> stringList = new List<string>();

            foreach(var item in arrImportante)
            {
                foreach(var item2 in item)
                {
                    stringList.Add(item2);
                }            
            }

            Console.WriteLine(stringList);
            Console.WriteLine(arrImportante);
            var textListAux = text.Split(null);

            var listAux2 = textListAux.Where(x => stringList.Any(y => y.ToLower() == x.ToLower())).ToList();
            string text2 = "";
            if (listAux2.Count != 0)
            {
                text2 = listAux2.Aggregate((a, b) => a + " " + b);
            }

            return GetFinalName(text2, brand); 
        } // public void GetModelNameByBrand(string text, string brand)


        public string GetFinalName(string text, string brand)
        {
            var aux = modelsByBrand.Where(x => x.Key.ToLower().Equals(brand.ToLower())).Select(y => y.Value);

            List<string> modelosPadrao = new List<string>();

            foreach (var item in aux)
            {
                foreach (var item2 in item)
                {
                    modelosPadrao.Add(item2);
                }

            }

            List<string> listAux4 = new List<string>();
            foreach (var item in modelosPadrao)
            {
                if (text.ToLower().Contains(item.ToLower()))
                {
                    listAux4.Add(item);
                }
            }

            Console.WriteLine(listAux4);

            var listOfValues = new Dictionary<string, int>();

            foreach (var item in listAux4)
            {
                listOfValues.Add(item, Compute(item, text));
            }


            var abc = listOfValues.OrderByDescending(key => key.Value);

            if (abc.Count() == 0)
            {
                text = "Padronizar Modelo - " + text;
            }
            else if (abc.Last().Value > 10)
            {
                text = "Modelo não identificado";
            }
            else
            {
                text = abc.Last().Key.Trim();
                text = text.TrimEnd();
                text = text.TrimStart();
            }
            return text;
        }

        /// <summary>
        /// Compute the distance between two strings.
        /// </summary>
        public int Compute(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Step 1
            if (n == 0)
            {
                return m;
            }

            if (m == 0)
            {
                return n;
            }

            // Step 2
            for (int i = 0; i <= n; d[i, 0] = i++)
            {
            }

            for (int j = 0; j <= m; d[0, j] = j++)
            {
            }

            // Step 3
            for (int i = 1; i <= n; i++)
            {
                //Step 4
                for (int j = 1; j <= m; j++)
                {
                    // Step 5
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;

                    // Step 6
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            // Step 7
            return d[n, m];
        }

        #region WriteExcelFile
        public void WriteExcelFile(string sheetName, string columnToRead, string columnToWrite, string path)
        {
            var app = new Excel.Application();
            app.Visible = false;
            app.DisplayAlerts = false;

            var excel = oWorkbook;      
            var table = excel.Worksheets[sheetName];

            var rowNo = table.UsedRange.Rows.Count;
            object[,] array = table.UsedRange.Value;

            string newValue = "";
            // assume that the 1st row = col name
            for(var rowIndex = 2; rowIndex <= rowNo; rowIndex++)
            {
                //GetModelNameByBrand("iphone 5s 16gb adas adoijsa motorola samsung teste ifone 4g", "APPLE");
                //IndiceColunaParaLer; IndiceColunaParaGravar; 2nd item = col position
                newValue = GetModelNameByBrand(table.Cells[rowIndex, IndiceColunaParaLer.Item2].Text, table.Cells[rowIndex, IndiceColunaFabricante2.Item2].Text); //read value
                table.Cells[rowIndex, IndiceColunaParaGravar.Item2] = newValue; //write new value
            }

            //foreach(var item in excel.Worksheets)
            //{
            //    app.Workbooks.Add(item);
            //}

            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(path))
                {
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(path);
                }

                table.SaveAs(path + "\\TesteCopiaExcel.xlsx");
              //  app.Quit();

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process to check/create folder failed: {0}", e.ToString());
            }
        }
        #endregion




    } // class DoThings

}
