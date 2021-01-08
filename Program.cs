
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Interop;



namespace LAB3
{
    class Program
    {
        static void Main(string[] args)
        {
            //Импорт из Excel в C#
            Application ObjExcel = new Application();
            var pathToFile = @"C:\Users\User\Desktop\Lab3.1.xlsm";
            Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Worksheet ObjWorkSheet;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            ObjExcel.Visible = true;

            //Список доступных функций
            Range xlRange = ObjWorkSheet.Range["A1","A4"];
            foreach (Range c in xlRange.Rows.Cells)
            {
                Console.WriteLine("Функция: " + c.Value);
            }

            //Запись выбранной функции в ячейку F2
            Console.WriteLine("Какую функцию посчитать?");
            int answer = Convert.ToInt32(Console.ReadLine());
            ObjWorkSheet.Cells[2, "F"] = answer;

            //Запись X
            Console.WriteLine("Выберите значение X");
            int x = Convert.ToInt32(Console.ReadLine());
            ObjWorkSheet.Cells[3, "F"] = x;

            //Таблица значений
            ObjWorkSheet.Cells[9, "A"] = "Таблица значений";
            ObjWorkSheet.Cells[10, "A"] = "X";
            ObjWorkSheet.Cells[10, "B"] = "Y";
            for (var i = 0; i <= x; i++)
            {
                
                ObjWorkSheet.Cells[3, "F"] = i;
                ObjWorkSheet.Cells[11+i, "A"] = i;
                ObjWorkSheet.Cells[11+i, "B"] = ObjWorkSheet.Cells[6, "F"];               
            }
            
            //Построение диаграммы
            ChartObjects xlCharts = (ChartObjects)ObjWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = xlCharts.Add(10, 80, 300, 250);
            Chart chartPage = myChart.Chart;
            myChart.Select();

            chartPage.ChartType = XlChartType.xlXYScatterSmooth;
            Application xla = new Application();
            SeriesCollection seriesCollection = chartPage.SeriesCollection();


            Series series1 = seriesCollection.NewSeries();
            series1.XValues = ObjWorkSheet.get_Range("A11", "A100"); 
            series1.Values = ObjWorkSheet.get_Range("B11", "B100");

            Console.Read();
        }
    }
}
