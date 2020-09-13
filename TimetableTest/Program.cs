using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TimetableTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            List<string> names = new List<string>();
            int sum = 0;
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Open(@"D:\Timetable.xlsx");
            Excel.Worksheet sheet = book.Worksheets[1];
            Excel.Range range = sheet.Range["A1", sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Address.ToString()];
            foreach (Excel.Range cell in range.Cells)
            {
                if (cell.Value2 != null && !names.Contains(cell.Value2.ToString()) && cell.Value2.ToString() != "ВЫХОДНЫЕ" && !cell.Value2.ToString().Contains(":") && !cell.Value2.ToString().Contains("4") && cell.Value2.ToString() != "Понедельник" && cell.Value2.ToString() != "Вторник" && cell.Value2.ToString() != "Среда" && cell.Value2.ToString() != "Четверг" && cell.Value2.ToString() != "Пятница" && cell.Value2.ToString() != "Суббота")
                {
                    names.Add(cell.Value2.ToString());
                }
            }
            Console.WriteLine();
            foreach (string name in names)
            {
                List<string> outs = new List<string>();
                int cost = 0;
                Console.WriteLine($"  {name}");
                foreach (Excel.Range cell in range.Cells)
                {
                    if (cell.Value2 != null)
                        if (cell.Value2.ToString() == name)
                        {
                            try
                            {
                                Excel.Range addressCell = sheet.Range[cell.Address.ToString().Remove(cell.Address.ToString().LastIndexOf("$")) + "2"];
                                Excel.Range timeCell = sheet.Cells[cell.Row, cell.Column - 1];
                                int addLength = 10 - DateTime.FromOADate(Convert.ToDouble(addressCell.Value2.ToString())).ToString("dddd").Length;
                                string addStr = null;
                                for (int i = 0; i <= addLength; i++)
                                    addStr += " ";
                                string output = $"{DateTime.FromOADate(Convert.ToDouble(addressCell.Value2.ToString())).ToString("d")}  {DateTime.FromOADate(Convert.ToDouble(addressCell.Value2.ToString())).ToString("dddd")} {addStr} {timeCell.Value2.ToString()}  {1800} руб.     ";
                                outs.Add(output);
                                cost += 1800;
                            }
                            catch { }
                        }
                }
                if (cost != 0)
                {
                    string temp;
                    for (int i = 0; i < outs.ToArray().Length - 1; i++)
                    {
                        for (int j = i + 1; j < outs.ToArray().Length; j++)
                        {
                            if (Convert.ToInt32(outs[i].Remove(outs[i].IndexOf("."))) > Convert.ToInt32(outs[j].Remove(outs[j].IndexOf("."))))
                            {
                                temp = outs[i];
                                outs[i] = outs[j];
                                outs[j] = temp;
                            }
                        }
                    }
                    foreach (string output in outs)
                        Console.WriteLine($"  {output}");
                    Console.WriteLine($"  Итого: {cost} руб.");
                    sum += cost;
                    Console.WriteLine("\n\n");
                }
            }
            Console.WriteLine($"Сумма: {sum} руб.");
            book = null;
            app.Quit();
            Console.ReadKey();
        }
    }
}
