using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public Excel.Worksheet Wsheet { get; set; }
        public Dictionary<int, string> Columns { get; set; }
        public Dictionary<string, int> Abas { get; set; }
        public List<KeyValuePair<string, string>> ValoresDePara { get; set; }
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

        public Dictionary<int, string> ValidarEntradasAbaDePara(string nomeAba)
        {
            Columns = new Dictionary<int, string>();

            var tst = oWorkbook.Worksheets[nomeAba];
            var colNo = tst.UsedRange.Columns.Count;
            object[,] array = tst.UsedRange.Value;
           
            for (var indice = 1; indice <= colNo; indice++)
            {
                if (array[1, indice] != null)
                {
                    Columns.Add(indice, array[1, indice].ToString());
                }
            }
            Console.WriteLine(Columns);

            return Columns;
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
                catch(Exception e)
                {

                }
            }



                return ValoresDePara;
        } // public Dictionary<string, string> CarregarValoresDePara

    } // class DoThings

}
