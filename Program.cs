
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
            dynamic excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                                //можно и так:
            //dynamic excelApp = new Application();
            excelApp.Workbooks.Open(@"C:\Users\User\Desktop\Lab3.1.xlsm");
            excelApp.Visible = true;
            dynamic workSheet = excelApp.ActiveSheet;

            //Импорт из Excel в C#
            //Application ObjExcel = new Application();
            //var pathToFile = @"C:\Users\User\Desktop\Lab3.1.xlsm";
            //Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Worksheet ObjWorkSheet;
            //ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            //ObjExcel.Visible = true;

            //Список доступных функций
            Range xlRange = excelApp.Range["A1", "A4"];
            foreach (Range c in xlRange.Rows.Cells)
            {
                Console.WriteLine("Функция: " + c.Value);
            }

            ////Запись выбранной функции в ячейку F2
            Console.WriteLine("Какую функцию посчитать?");
            int answer = Convert.ToInt32(Console.ReadLine());
            workSheet.Cells[2, "F"] = answer;

            ////Запись X
            Console.WriteLine("Выберите значение X");
            int x = Convert.ToInt32(Console.ReadLine());
            workSheet.Cells[3, "F"] = x;

            ////Таблица значений
            workSheet.Cells[9, "A"] = "Таблица значений";
            workSheet.Cells[10, "A"] = "X";
            workSheet.Cells[10, "B"] = "Y";
            for (var i = 0; i <= x; i++)
            {

                workSheet.Cells[3, "F"] = i;
                workSheet.Cells[11 + i, "A"] = i;
                workSheet.Cells[11 + i, "B"] = workSheet.Cells[6, "F"];
            }
            
            

            ////Построение диаграммы
            ChartObjects xlCharts = (ChartObjects)workSheet.ChartObjects(Type.Missing);
            ChartObject myChart = xlCharts.Add(10, 80, 300, 250);
            Chart chartPage = myChart.Chart;
            myChart.Select();

            chartPage.ChartType = XlChartType.xlXYScatterSmooth;
            Application xla = new Application();
            SeriesCollection seriesCollection = chartPage.SeriesCollection();


            Series series1 = seriesCollection.NewSeries();
            series1.XValues = excelApp.get_Range("A11", "A100");
            series1.Values = excelApp.get_Range("B11", "B100");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.Read();
        }
    }
}
