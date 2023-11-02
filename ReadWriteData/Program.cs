using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace ReadWriteData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string sourceFilePath = @"C:\Users\Nimap\Downloads\week 38\Presale - CJSC 09-25-2023.xlsx";
            string targetFilePath = @"C:\Users\Nimap\Downloads\week 38\Week 38 Sales 2023-09-18.xlsm";
            var lastNbr = int.MinValue;
            using (var excelpackage = new ExcelPackage(new FileInfo(targetFilePath)))
            {
                using (var sourceFile = new ExcelPackage(new FileInfo(sourceFilePath)))
                {
                    
                    var worksheetNames = excelpackage.Workbook.Worksheets.Where(x => x.Name.StartsWith("Week")).ToList();
                    foreach (var item in worksheetNames)
                    {
                        var weekNbr = Convert.ToInt32(item.Name.Split(' ')[1]);
                        lastNbr = Math.Max(lastNbr, weekNbr);
                    }
                    var presalesWorksheet = sourceFile.Workbook.Worksheets["Presale"];
                    ExcelCalculationOption calculationOption = new ExcelCalculationOption();
                    calculationOption.AllowCircularReferences = true;
                    presalesWorksheet.Calculate(calculationOption);
                    presalesWorksheet.ClearFormulas();
                    sourceFile.Save();
                    var week1Worksheet = excelpackage.Workbook.Worksheets.Add($"Week {lastNbr + 1}", presalesWorksheet);
                    
                }
                excelpackage.Save();
            }
            // Create a new Excel Application
            //Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            //// Open the source Excel file
            //Excel.Workbook sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
            //Excel.Workbook targetWorkbook = excelApp.Workbooks.Open(targetFilePath);

            //// Access the third worksheet (sheet[3]) in the source file
            //int sourceSheetIndex = 3;
            //Excel.Worksheet sourceWorksheet = (Excel.Worksheet)sourceWorkbook.Sheets[sourceSheetIndex];

            //// Create a new Excel Workbook for the target file
            //Excel.Worksheet targetWorsheet = targetWorkbook.Worksheets.Add();

            ////List<string> workSheetNames = new List<string>();
            //foreach (Worksheet item in targetWorkbook.Worksheets)
            //{
            //    if (item.Name.Equals("Week"))
            //    {
            //        //workSheetNames.Add(item.Name);\
            //        var weekNbr = Convert.ToInt32(item.Name.Split(' ')[1]);
            //        Math.Max(lastNbr, weekNbr);
            //    }
            //}

            //try
            //{
            //    // Copy the source worksheet to the target workbook
            //    sourceWorksheet.Copy(Type.Missing, targetWorkbook.Sheets[targetWorkbook.Sheets.Count]);

            //    // Save the target workbook with the copied worksheet
            //    targetWorkbook.SaveAs(targetFilePath);

            //    // Close the source and target workbooks
            //    sourceWorkbook.Close();
            //    targetWorkbook.Close();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("An error occurred: " + ex.Message);
            //}
            //finally
            //{
            //    // Close the Excel application
            //    excelApp.Quit();
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWorkbook);
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //}
        }
    }


}
