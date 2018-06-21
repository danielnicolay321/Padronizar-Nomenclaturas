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

        public List<KeyValuePair<string, Dictionary<int, string>>> DeParaColumns = new List<KeyValuePair<string, Dictionary<int, string>>>();

      //  public Dictionary<int, string> DeParaColumns { get; set; }
        public Dictionary<int, string> ReadAndWriteColumns { get; set; }
        public Dictionary<string, int> Abas { get; set; }
        public List<KeyValuePair<string, string>> ValoresDePara { get; set; }
        public List<KeyValuePair<string, List<string>>> wordListByBrand { get; set; }
        public List<KeyValuePair<string, List<string>>> modelsByBrand { get; set; }
        public Sheets sheets { get; set; }

        Tuple<int, int> IndiceColunaModeloPadrao = new Tuple<int, int>(0, 0);
        Tuple<int, int> IndiceColunaFabricante = new Tuple<int, int>(0, 0);

        public DoThings(string nomeItem)
        {
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

            var tst = oWorkbook.Worksheets[nomeAba];
            var colNo = tst.UsedRange.Columns.Count;
            var rowNo = tst.UsedRange.Rows.Count;
            object[,] array = tst.UsedRange.Value;

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
            GetModelNameByBrand("iphone 5s 16gb adas adoijsa motorola samsung teste ifone 4g", "APPLE");
        } // public void SepararModelosPorFabricante(List<KeyValuePair<string, List<string>>> modelsByBrand)


        public void GetModelNameByBrand(string text, string brand)
        {
            var arrImportante = wordListByBrand.Where(x => x.Key.ToLower().Equals(brand.ToLower())).Select(y => y.Value);

            List<string> tsts = new List<string>();

            foreach(var item in arrImportante)
            {
                foreach(var item2 in item)
                {
                    tsts.Add(item2);
                }
                
            }

            Console.WriteLine(tsts);
            Console.WriteLine(arrImportante);
            var textListAux = text.Split(null);

            var listAux2 = textListAux.Where(x => tsts.Any(y => y.ToLower() == x.ToLower())).ToList();
            GetFinalName(listAux2.Aggregate((a, b) => a + " " + b), brand); 
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

            var teste = new Dictionary<string, int>();

            foreach (var item in listAux4)
            {
                teste.Add(item, Compute(item, text));
            }


            var abc = teste.OrderByDescending(key => key.Value);

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




        //public static System.Data.DataTable ExportToExcel(Excel.Worksheet oWorksheet)
        //{
        //    System.Data.DataTable table = new System.Data.DataTable();

        //    table.Columns.Add("SEMANA_MES", typeof(string)); // 1
        //    table.Columns.Add("SEMANA_ANO", typeof(string)); // 2
        //    table.Columns.Add("MODELO_PADRONIZADO", typeof(string));// 3
        //    table.Columns.Add("ORIGEM", typeof(string));  // 4
        //    table.Columns.Add("DATA", typeof(string)); // 5
        //    table.Columns.Add("FABRICANTE", typeof(string));
        //    table.Columns.Add("MODELO_ORIGINAL", typeof(string)); // 6
        //    table.Columns.Add("SITUACAO_DO_APARELHO", typeof(string)); // 8
        //    table.Columns.Add("TIPO_NEGOCIACAO_CLIENTE", typeof(string)); // 9
        //    table.Columns.Add("VALOR", typeof(string)); // 10
        //    table.Columns.Add("CONDICAO_DE_PAGAMENTO", typeof(string));// 11
        //    table.Columns.Add("URL_FONTE", typeof(string)); // 12

        //    //nr de colunas da planilha carregada
        //    int colNo = oWorksheet.UsedRange.Columns.Count;
        //    //nr de linhas da planilha carregada
        //    int rowNo = oWorksheet.UsedRange.Rows.Count;
        //    // read the value into an array.
        //    object[,] array = oWorksheet.UsedRange.Value;

        //    string modeloOriginal = "";
        //    string modeloPadronizado = "";
        //    string urlfonte = "";
        //    //    string valor = "";
        //    string condicaoPagamento = "";
        //    string situacaoAparelho = "";
        //    string fabricante = "";
        //    string orig = "";
        //    string semanaMes = "";
        //    string semanaAno = "";

        //    string ORIGEM = "";
        //    string SEMANA_MES = "";
        //    string SEMANA_ANO = "";
        //    string FABRICANTE = "";
        //    string MODELO_ORIGINAL = "";
        //    string MODELO_PADRONIZADO = "";
        //    string SITUACAO_DO_APARELHO = "";
        //    string TIPO_NEGOCIACAO_CLIENTE = "";
        //    string VALOR = "";
        //    //string VALOR_DE_COMPRA = "";
        //    //string VALOR_DE_VENDA = "";
        //    string CONDICAO_DE_PAGAMENTO = "";
        //    string URL_FONTE = "";
        //    string DATA = "";


        //    // j = NR DE LINHAS 
        //    for (int j = 1; j <= rowNo; j++)
        //    {
        //        string modeloPadronizadoSemLixo = "";

        //        int contAux = j + 1;
        //        if (contAux > rowNo)
        //        {
        //            break;
        //        }
        //        try
        //        {
        //            FABRICANTE = array[contAux, 6].ToString();
        //            var desconsiderar = IdentificarFabricante(FABRICANTE);
        //            if (!desconsiderar)
        //            {
        //                continue;
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            FABRICANTE = "";
        //        }


        //        try
        //        {
        //            SEMANA_MES = array[contAux, 1].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            SEMANA_MES = "";
        //        }

        //        try
        //        {

        //            SEMANA_ANO = array[contAux, 2].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            SEMANA_ANO = "";
        //        }
        //        try
        //        {
        //            MODELO_PADRONIZADO = SetBrandParameters(array[contAux, 6].ToString(), array[contAux, 7].ToString(), array[contAux, 12].ToString());
        //            MODELO_PADRONIZADO = MODELO_PADRONIZADO.Trim();
        //        }
        //        catch (Exception e)
        //        {
        //            MODELO_PADRONIZADO = "";
        //        }

        //        try
        //        {
        //            ORIGEM = array[contAux, 4].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            ORIGEM = "";
        //        }

        //        try
        //        {
        //            DATA = array[contAux, 5].ToString();
        //            var spt = DATA.Split(null);

        //            DATA = spt[0];
        //            var abc = DateTime.Parse(DATA).Date;
        //            var abc2 = abc.ToString("dd/MM/yyyy");
        //            Console.WriteLine(abc2);
        //        }
        //        catch (Exception e)
        //        {
        //            DATA = "";
        //        }

        //        try
        //        {
        //            MODELO_ORIGINAL = array[contAux, 7].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            MODELO_ORIGINAL = "";
        //        }
        //        try
        //        {
        //            SITUACAO_DO_APARELHO = array[contAux, 8].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            SITUACAO_DO_APARELHO = "";
        //        }
        //        try
        //        {
        //            TIPO_NEGOCIACAO_CLIENTE = array[contAux, 9].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            TIPO_NEGOCIACAO_CLIENTE = "";
        //        }

        //        try
        //        {
        //            VALOR = array[contAux, 10].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            VALOR = "";
        //        }

        //        try
        //        {
        //            CONDICAO_DE_PAGAMENTO = array[contAux, 11].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            CONDICAO_DE_PAGAMENTO = "";
        //        }
        //        try
        //        {
        //            URL_FONTE = array[contAux, 12].ToString();
        //        }
        //        catch (Exception e)
        //        {
        //            URL_FONTE = "";
        //        }


        //        table.Rows.Add(
        //            SEMANA_MES,
        //            SEMANA_ANO,
        //            MODELO_PADRONIZADO, //5 MODELO_PADRONIZADO (TRATAMENTO)
        //            ORIGEM, // 1 ORIGEM
        //            DATA,  //2  DATA
        //            FABRICANTE, //3 FABRICANTE
        //            MODELO_ORIGINAL,//4 MODELO_ORIGINAL                  
        //            SITUACAO_DO_APARELHO,//6 SITUACAO_DO_APARELHO
        //            TIPO_NEGOCIACAO_CLIENTE,//7 TIPO_NEGOCIACAO
        //            VALOR, //8 VALOR
        //            CONDICAO_DE_PAGAMENTO, //11 CONDICAO_DE_PAGAMENTO
        //            URL_FONTE);//12 URL_FONTE

        //    } // for int j


        //    return table;
        //}




        #region WriteExcelFile
        public static void WriteExcelFile(System.Data.DataTable dataTable)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            worKbooK = excel.Workbooks.Add(Type.Missing);
            worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
            int rowcount = 2;

            foreach (DataRow datarow in dataTable.Rows)
            {
                rowcount += 1;
                for (int i = 1; i <= dataTable.Columns.Count; i++)
                {

                    if (rowcount == 3)
                    {
                        worKsheeT.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;

                    }

                    worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                    if (rowcount > 3)
                    {
                        if (i == dataTable.Columns.Count)
                        {
                            if (rowcount % 2 == 0)
                            {
                                celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, dataTable.Columns.Count]];

                            }
                        }
                    } // if (rowcount > 3)
                } // for (int i = 1; i <= tst.Columns.Count; i++)
            } // foreach (DataRow datarow in tst.Rows)


            celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, dataTable.Columns.Count]];
            celLrangE.EntireColumn.AutoFit();
            celLrangE.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            celLrangE.Font.Size = 8;
            celLrangE.Font.FontStyle = "Calibri";

            Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, dataTable.Columns.Count]];

            var path = @"C:\Relatorio";

            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(path))
                {
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(path);
                }

                worKbooK.SaveAs(path + "\\planilha_semanal_Robo.xlsx");
                worKbooK.Close();

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process to check/create folder failed: {0}", e.ToString());
            }

            excel.Quit();
        }
        #endregion




    } // class DoThings

}
