using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;//biblioteca para manipular o Excel
using _Excel = Microsoft.Office.Interop.Excel;//atribui uma "instancia" para manipular o Excel 

namespace HealthScore
{
    class Excel
    {
        string pathWorkbook = "";
        _Application excel = new _Excel.Application();
        Workbook modelWorkbook;
        Worksheet someWorkshhet;

        public Excel(string pathWorkbook, int worksheetNumber) //abre o excel de acordo com os parametros
        {
            this.pathWorkbook = pathWorkbook;
            modelWorkbook = excel.Workbooks.Open(pathWorkbook);
            someWorkshhet = modelWorkbook.Worksheets[worksheetNumber];
        }

        public string ReadValueCell(int row, int column) //le os valores de acordo de cada célula
        {
            if (someWorkshhet.Cells[row, column].Value2 != null)
                return (someWorkshhet.Cells[row, column].Value2).ToString();
            else
                return "";
        }

        public void SetHeaderValue(int column, string header) //grava os valores do cabeçalho
        {
            int row = 2;
            column += 2;

            someWorkshhet.Cells[row, column].Value2 = header;
        }

        public void SetValueCell(int row, int id, int healthScore, 
                                 int pneumonia, int breastCancer, int hipFracture,
                                 int parkinsonsDisease, int death) //grava os valores processados
        {
            int column = 1;

            someWorkshhet.Cells[row, column].Value2 = id;
            someWorkshhet.Cells[row, (column+1)].Value2 = healthScore;
            someWorkshhet.Cells[row, (column+2)].Value2 = pneumonia;
            someWorkshhet.Cells[row, (column+3)].Value2 = breastCancer;
            someWorkshhet.Cells[row, (column+4)].Value2 = hipFracture;
            someWorkshhet.Cells[row, (column+5)].Value2 = parkinsonsDisease;
            someWorkshhet.Cells[row, (column+6)].Value2 = death;

        }

        public void Save() //salva a planilha
        {
            modelWorkbook.Save();          
        }

        public void Close() //fecha a planilha
        {
            modelWorkbook.Close();
        }
    }
}
