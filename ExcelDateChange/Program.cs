using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelDateChange
{
    class Program
    {
        const string signatureFileName = "signatures.xlsx";
        const string excelFilesDirectory = "excelExample";
        static string currentDirectory = Directory.GetCurrentDirectory();
        static string[] notPricesSlices = new string[] { "ФЕР", "ФССЦ", "фссц", "ФСЭМ" };

        static void Main(string[] args)
        {
            Console.WriteLine("Press number:");
            Console.WriteLine("1 - to change value of one cell into all excel files in subdirectory \\{0}", excelFilesDirectory);
            Console.WriteLine("2- to add block of signatures from file \"{0}\" into the end of all excel files in subdirectory \\{1}", signatureFileName, excelFilesDirectory);
            Console.WriteLine("3 - to get all price position from all excel files in subdirectory \\{0}", excelFilesDirectory);
            Console.WriteLine();
            string option = "";
            bool b = true;
            while (b)
            {
                Console.Write("Please, print number of option, what you want to do: ");
                option = Console.ReadLine();
                switch (option)
                {
                    case "1":
                        ChangeCell();
                        b = false;
                        break;
                    case "2":
                        AddSignBloc();
                        b = false;
                        break;
                    case "3":
                        FindPrices();
                        b = false;
                        break;
                    default:
                        Console.WriteLine("Please, print correct option");
                        break;
                }
            }
            Console.WriteLine("All done!");
            Console.ReadLine();
            
        }

        static void AddSignBloc()
        {
            DirectoryInfo di = new DirectoryInfo(currentDirectory+"\\"+excelFilesDirectory);
            List<FileInfo> fil = di.GetFiles("*.xlsx").ToList<FileInfo>();
            Application app = new Application();
            //app.Visible = true;
            Workbook signWorkbook = app.Workbooks.Open(currentDirectory+"\\"+signatureFileName, ReadOnly: true);
            Worksheet signWorksheet = signWorkbook.Sheets[1];
            int signRowCount = signWorksheet.Rows.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("#######################");
                Console.WriteLine("Work with {0}", fi.Name);
                Workbook wb = app.Workbooks.Open(fi.FullName, ReadOnly: false);
                Worksheet ws = wb.Worksheets[1];
                int lastRow = ws.Rows.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                signWorksheet.Rows["1:" + signRowCount.ToString()].Copy();
                ws.Rows[lastRow+3].PasteSpecial();
                wb.Save();
                wb.Close();
                Console.WriteLine("Done!");

            }
        }

        static void ChangeCell()
        {
            Console.Write("Print row of your cell: ");
            int row = Convert.ToInt32(Console.ReadLine());
            Console.Write("Print column of your cell: ");
            int column = Convert.ToInt32(Console.ReadLine());
            Console.Write("Print value of your cell: ");
            string value = Console.ReadLine();

            DirectoryInfo di = new DirectoryInfo(currentDirectory+"\\"+excelFilesDirectory);
            List<FileInfo> fil = di.GetFiles("*.xlsx").ToList<FileInfo>();
            Application app = new Application();
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("#######################");
                Console.WriteLine("Work with {0}", fi.Name);
                Workbook wb = app.Workbooks.Open(fi.FullName, ReadOnly: false);
                Worksheet ws = wb.Worksheets[1];
                ws.Cells[row, column].Value = value;
                wb.Save();
                wb.Close();
                Console.WriteLine("Done!");
            }
        }

        static void FindPrices()
        {
            DirectoryInfo di = new DirectoryInfo(currentDirectory + "\\" + excelFilesDirectory);
            List<FileInfo> fil = di.GetFiles("*.xlsx").ToList<FileInfo>();
            Application app = new Application();
            app.Visible = true;
            Workbook resultWorkbook = app.Workbooks.Add();
            Worksheet resultWorksheet = resultWorkbook.Worksheets.Add();
            int resultRow = 1;
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("#######################");
                Console.WriteLine("Work with {0}", fi.Name);
                Workbook wb = app.Workbooks.Open(fi.FullName, ReadOnly: false);
                Worksheet ws = wb.Worksheets[1];
                resultWorksheet.Cells[resultRow++, 1].Value = fi.Name;
                int currentRow = 30;
                int lastRow = ws.Rows.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                while (currentRow < lastRow)
                {
                    if (ws.Cells[currentRow, 3].Font.Bold == true)
                    {
                        string code = ws.Cells[currentRow, 3].Value.ToString(); ;
                        /*if ((code != null) && (code.IndexOf("ФЕР") == -1) && (code.IndexOf("ФССЦ") == -1)
                            && (code.IndexOf("фссц") == -1) && (code.IndexOf("ФСЭМ") == -1))
                        {
                            ws.Rows[currentRow].Copy();
                            resultWorksheet.Rows[resultRow++].PasteSpecial();
                            //resultWorksheet.Rows[resultRow++] = ws.Rows[currentRow];
                        }*/
                        if (code != null)
                        {
                            bool isPrices = true;
                            foreach (string slice in notPricesSlices)
                            {
                                if (code.IndexOf(slice) != -1) isPrices = false;
                            }
                            if (isPrices)
                            {
                                ws.Rows[currentRow].Copy();
                                resultWorksheet.Rows[resultRow++].PasteSpecial();
                            }
                        }
                    }
                    currentRow++;
                }
                wb.Application.CutCopyMode = (XlCutCopyMode)0;
                wb.Close();
                Console.WriteLine("Done!");
            }
        }
    }
}
