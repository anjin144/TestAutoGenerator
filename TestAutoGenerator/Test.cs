using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestAutoGenerator
{
    public static class Test
    {
        public static void Search()
        {
            Excel.Application excelApplication = new Excel.Application();
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(@"C:\Users\anjin\Documents\Visual Studio 2015\Projects\TestAutoGenerator\TestAutoGenerator\bin\Debug\test1.xlsx");
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);


            // var failure0 = AppManager.GetFailureByName(exigences[1].FailureName, ConfigurationManager.AppSettings["InputFile2_SheetName_Activate"]);
            // AppManager.ProcessFailure(templateFilePath, outputFilePath, exigences[1], failure0);
            //// return;

            // failure0 = AppManager.GetFailureByName(exigences[2].FailureName, ConfigurationManager.AppSettings["InputFile2_SheetName_Activate"]);
            // AppManager.ProcessFailure(templateFilePath, outputFilePath, exigences[2], failure0);

            /* //working filter
            Excel.Range range = srcworkSheet.UsedRange;
            range.AutoFilter(1, "AAA", Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            var filteredRange = range.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            Console.WriteLine(filteredRange.Rows.Count);
            // end working filter */

            string rangeTo = "A1:A" + srcworkSheet.UsedRange.Rows.Count;
            var range = srcworkSheet.Range[rangeTo];

            //var rowsCol0 = range.Find("AAA", Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
            //               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, false, false);
            range.Find(What: "AAA", LookIn: Excel.XlFindLookIn.xlValues,
            LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns);

            Excel.Range start = srcworkSheet.Range["A1"];
            //srcworkSheet.Cells.Find("A",
            //                    Excel.XlLookAt.xlPart,
            //                    Excel.XlSearchOrder.xlByColumns,
            //                    Excel.XlSearchDirection.xlNext);

            HashSet<int> matches = new HashSet<int>();

            Excel.Range next = start;

            while (true)
            {
                next = range.FindNext(next.Offset[1, 0]);
                //next = range.FindNext("AAA");
                if (!matches.Add(next.Row))
                    break;
            }

            //Console.WriteLine(rowsCol0.Rows.Value);
            //Console.WriteLine(rngResult.Rows.Value);

            srcworkBook.Close(true);
            //Kill excelapp
            excelApplication.Quit();
        }


    }
}
