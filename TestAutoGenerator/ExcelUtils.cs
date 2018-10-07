using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestAutoGenerator
{
    public static class ExcelUtils
    {
        #region NPOI methods
        public static int GetColumnIndexByText(string filePath, string text)
        {
            int result = 0;

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            var rowHeaderIndex = int.Parse(ConfigurationManager.AppSettings["InputFile1_Header_RowIndex"]);
            ISheet sheet = hssfwb.GetSheet(ConfigurationManager.AppSettings["InputFile1_SheetName"]);
            if (sheet == null)
                return -1;

            var cell = GetColumnByContainingText(sheet.GetRow(rowHeaderIndex), text);
            if (cell != null)
                result = cell.ColumnIndex;

            return result;
        }

        public static int GetColumnIndexByText(ISheet sheet, int rowHeaderIndex, string text)
        {
            int result = -1;

            var cell = GetColumnByContainingText(sheet.GetRow(rowHeaderIndex), text);
            if (cell != null)
                result = cell.ColumnIndex;

            return result;
        }

        private static ICell GetColumnByContainingText(IRow rowHeading, string text)
        {
            if (rowHeading != null)
            {
                return rowHeading.Cells.Find(c => (c.CellType == CellType.String || c.CellType == CellType.Unknown) && c.StringCellValue != null && c.StringCellValue.Contains(text));
            }
            else
                return null;
        }

        public static Dictionary<int, string> GetColumnsAfterIndex(string filePath, int startingIndex, int rowIndex)
        {
            var result = new Dictionary<int, string>();

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet(ConfigurationManager.AppSettings["InputFile1_SheetName"]);
            IRow row = sheet.GetRow(rowIndex);  //2

            for (int i = startingIndex; i < row.Cells.Count; i++)
            {
                var colVal = row.Cells[i].StringCellValue;
                result.Add(i, colVal);
            }

            return result;
        } 

        public static List<IRow> GetRowsByColumnIndexAndCellValue(string filePath, string sheetName, int colIndex, string searchBy)
        {
            List<IRow> lstResult = new List<IRow>();

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet(sheetName);            
            
            foreach (IRow row in sheet)
            {
                var cell = row.GetCell(colIndex);
                if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != null && cell.StringCellValue.ToUpper() == searchBy.ToUpper())
                    lstResult.Add(row);
            }

            return lstResult;
        }

        public static List<IRow> GetRowsByColumnIndexAndCellValue(ISheet sheet, int colIndex, string searchBy)
        {
            List<IRow> lstResult = new List<IRow>();

            foreach (IRow row in sheet)
            {
                var cell = row.GetCell(colIndex);
                if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != null && cell.StringCellValue.ToUpper() == searchBy.ToUpper())
                    lstResult.Add(row);
            }

            return lstResult;
        }

        public static void WriteFailureDataToExcel(Excel.Application excelApplication, string templateFilePath, string fileTarget, ExigenceDetails exigence, Failure failure, Homologation homologation, int indexRecord)
        {
            try
            {
                Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(templateFilePath);
                Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

                // modify template
                FillInTemplate(srcworkSheet, exigence, failure, homologation, indexRecord);

                Excel.Workbook destworkBook = excelApplication.Workbooks.Open(fileTarget, 0, false);
                Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

                // determine rangeTo
                int sourceLinesNb = srcworkSheet.UsedRange.Rows.Count;
                int lastRow = destworkSheet.UsedRange.Rows.Count;
                //if (lastRow < 10)
                    //lastRow += 1;

                string rangeTo = "A" + (lastRow + 1) + ":H" + (lastRow + 1 + sourceLinesNb); //(lastRow + 36);
                Console.WriteLine("rangeTo -> " + rangeTo);

                Excel.Range from = srcworkSheet.UsedRange;
                Excel.Range to = destworkSheet.Range[rangeTo];

                // if you use 2 instances of excel, this will not work
                from.Copy(to);

                destworkBook.SaveAs(fileTarget);

                destworkBook.Close(true, null, null);
                srcworkBook.Close(false, null, null);
            }
            catch (Exception ex)
            {
                Logger.Log("Error in processing the exigence: " + exigence.FailureName + ". Details: " + ex.Message + " -- " + ex.StackTrace);
            }
        }
        #endregion

        #region Excel lib methods

        /// <summary>
        /// Clear the output file (remove all lines copied from template - only header remains)
        /// Keep only one sheet for fuel type (diesel or gasoline)
        /// </summary>
        /// <param name="path"></param>
        /// <param name="removeSheet"></param>
        public static void InitiateOutputFile(string path, string removeSheet)
        {
            var excelApplication = new Excel.Application();
            excelApplication.Application.DisplayAlerts = false;

            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(path);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            var indexHeader = GetHeaderTemplateIndex(srcworkSheet);
            string rangeTo = "A" + (indexHeader + 1) + ":H" + (srcworkSheet.UsedRange.Rows.Count + 1);
            Excel.Range range = srcworkSheet.Range[rangeTo];
            range.EntireRow.Delete(Type.Missing);

            // remove one fuel type sheet
            var worksheet = srcworkBook.Sheets[removeSheet];
            worksheet.Delete();

            var worksheetFuel = srcworkBook.Worksheets.get_Item(2);
            worksheetFuel.Name = ConfigurationManager.AppSettings["OutputFile_SheetFuelTypeName"];

            srcworkBook.SaveAs(path);
            srcworkBook.Close(false, null, null);
            excelApplication.Quit();
        }

        public static void FinishOutputFile(string path)
        {
            var excelApplication = new Excel.Application();
            excelApplication.Application.DisplayAlerts = false;

            try
            {
                Excel.Workbook wb = excelApplication.Workbooks.Open(path);
                Excel.Worksheet ws = wb.Worksheets.get_Item(1);

                ws.Rows[ws.UsedRange.Rows.Count + 2].Insert();
                ws.Cells[ws.UsedRange.Rows.Count + 1, 1] = "\\End";

                wb.SaveAs(path);
                wb.Close(false, null, null);
            }
            catch (Exception ex)
            {
                Logger.Log("Error in FinishOutputFile: " + ex.Message);
            }
            finally
            {
                excelApplication.Quit();
            }            
        }

        private static int GetHeaderTemplateIndex(Excel.Worksheet srcworkSheet)
        {
            int lastRow = srcworkSheet.UsedRange.Rows.Count;
            var rowHeader = GetRowsIndexesByText(srcworkSheet.Range["A1"], srcworkSheet.Range["A1:A" + lastRow], "Chapter");
            var headerIndex = 1;
            if (rowHeader.Count > 0)
                headerIndex = rowHeader[rowHeader.Count - 1];

            return headerIndex;
        }

        public static Excel.Worksheet FillInTemplate(Excel.Worksheet sheet, ExigenceDetails exigence, Failure failure, Homologation homologation, int indexRecord)
        {
            #region remove template header
            var indexHeader = GetHeaderTemplateIndex(sheet);
            string rangeToHeader = "A1:A" + (indexHeader);
            Console.WriteLine("range header: " + rangeToHeader);
            Excel.Range rangeHeader = sheet.Range[rangeToHeader];
            rangeHeader.EntireRow.Delete(Type.Missing);
            #endregion

            string rangeB_To = null;
            string rangeTo = "A1:A" + sheet.UsedRange.Rows.Count;
            var rangeA = sheet.Range[rangeTo];

            #region Replace column A - Failure_name
            var rowsC1 = GetRowsIndexesByText(sheet.Range["A1"], rangeA, "Failure_name");
            if (rowsC1 != null)
            {
                foreach (var row in rowsC1)
                {
                    sheet.Cells[row, 1] = failure.Chapter;
                }
            }
            #endregion

            #region Replace column F - Failure_name & TdH
            rangeTo = rangeTo.Replace('A', 'F');
            var rowsC6 = GetRowsIndexesByText(sheet.Range["F1"], sheet.Range[rangeTo], "Failure_name");
            if (rowsC6 != null)
            {
                Console.WriteLine("rows count c6: " + rowsC6.Count);
                foreach (var row in rowsC6)
                {
                    sheet.Cells[row, 6] = failure.Chapter;
                }
            }

            rowsC6 = GetRowsIndexesByText(sheet.Range["F1"], sheet.Range[rangeTo], "Read from TdH");
            if (rowsC6 != null)
            {
                Console.WriteLine("rows count c6 tdh: " + rowsC6.Count);
                var rowIncrement = 0;
                foreach (var row in rowsC6)
                {
                    // modify row (if new lines were added/removed)
                    var index = row + rowIncrement;

                    if (homologation == null || homologation.FailingDiagCode == null)   // remove homologation line
                    {
                        Excel.Range rng = sheet.Range["F" + index];
                        rng.EntireRow.Delete(Type.Missing);
                        rowIncrement--;
                        continue;
                    }
                    
                    // fill in homologation line
                    sheet.Cells[index, 6] = homologation.FailingDiagCode_Transformed;
                }
            }
            #endregion

            #region Replace column G - Failure_name, Vbx_wal_1_req, Vbx_wal_2_req, Vbx_lih_typ_4, Vbx_mil_ecu_req, Vbx_mil_flsh_req
            rangeTo = "G1:G" + sheet.UsedRange.Rows.Count;
            var rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Vbx_wal_1_req");
            if (rowsC7 != null)
            {
                foreach (var row in rowsC7)
                {
                    if (sheet.Cells[row, 8].Value != null && Convert.ToString(sheet.Cells[row, 8].Value).Trim() == "Read from Exigence SdF")
                    {
                        if (exigence.G1) //if (exigence.G1.ToUpper() == "YES" || exigence.G1.ToUpper() == "OUI")
                            sheet.Cells[row, 8] = 1;
                        else
                            sheet.Cells[row, 8] = 0;
                    }
                }
            }

            rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Vbx_wal_2_req");
            if (rowsC7 != null)
            {
                foreach (var row in rowsC7)
                {
                    if (sheet.Cells[row, 8].Value != null && Convert.ToString(sheet.Cells[row, 8].Value).Trim() == "Read from Exigence SdF")
                    {
                        if (exigence.G2) //if (exigence.G2.ToUpper() == "YES" || exigence.G2.ToUpper() == "OUI")
                            sheet.Cells[row, 8] = 1;
                        else
                            sheet.Cells[row, 8] = 0;
                    }
                }
            }

            rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Vbx_lih_typ_4");
            if (rowsC7 != null)
            {
                foreach (var row in rowsC7)
                {
                    if (sheet.Cells[row, 8].Value != null && Convert.ToString(sheet.Cells[row, 8].Value).Trim() == "Read from Exigence SdF")
                    {
                        if (exigence.GEE) //if(exigence.GEE.ToUpper() == "NON" || exigence.GEE.ToUpper() == "NO")
                            sheet.Cells[row, 8] = 0;
                        else
                            sheet.Cells[row, 8] = 1;
                    }
                }
            }

            rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Failure_name");
            if (rowsC7 != null)
            {
                foreach (var row in rowsC7)
                {
                    var aux = failure.Chapter.Remove(0, 8);
                    aux = Convert.ToString(sheet.Cells[row, 7].Value).Replace("failure_name", aux);

                    sheet.Cells[row, 7] = aux;
                }
            }

            rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Vbx_mil_ecu_req");
            if (rowsC7 != null)
            {
                Console.WriteLine("rows count c7 Vbx_mil_ecu_req: " + rowsC7.Count);
                var rowIncrement = 0;
                foreach (var row in rowsC7)
                {
                    // modify row (if new lines were added/removed)
                    var index = row + rowIncrement;

                    if (homologation == null || homologation.FailingDiagCode == null)   // remove homologation line
                    {
                        Excel.Range rng = sheet.Range["G" + index];
                        rng.EntireRow.Delete(Type.Missing);
                        rowIncrement--;
                        continue;
                    }

                    // fill in homologation line
                    if (homologation.MIL.ToUpper().Trim() == "YES" || homologation.MIL.ToUpper().Trim() == "OUI")
                        sheet.Cells[index, 8] = 1;
                    else
                        sheet.Cells[index, 8] = 0;
                }
            }

            rangeTo = "G1:G" + sheet.UsedRange.Rows.Count;
            rowsC7 = GetRowsIndexesByText(sheet.Range["G1"], sheet.Range[rangeTo], "Vbx_mil_flsh_req");
            if (rowsC7 != null)
            {
                Console.WriteLine("rows count c7 Vbx_mil_flsh_req: " + rowsC7.Count);
                var rowIncrement = 0;
                foreach (var row in rowsC7)
                {
                    // modify row (if new lines were added/removed)
                    var index = row + rowIncrement;

                    if (homologation == null || homologation.FailingDiagCode == null)   // remove homologation line
                    {
                        Excel.Range rng = sheet.Range["G" + index];
                        rng.EntireRow.Delete(Type.Missing);
                        rowIncrement--;
                        continue;
                    }

                    // fill in homologation line
                    if (homologation.MIL.ToUpper().Trim().Contains("CLIGNOTANT") || homologation.MIL.ToUpper().Trim().Contains("FLASHING"))
                        sheet.Cells[index, 8] = 1;
                    else
                        sheet.Cells[index, 8] = 0;
                }
            }
            #endregion

            #region Replace column H - Failure_name
            /*rangeTo = "H1:H" + sheet.UsedRange.Rows.Count;
            //rangeTo = rangeTo.Replace('G', 'H');
            var rowsC8 = GetRowsIndexesByText(sheet.Range["H1"], sheet.Range[rangeTo], "failure_name");
            if (rowsC8 != null)
            {
                foreach (var row in rowsC8)
                {
                    sheet.Cells[row, 8] = failure.Chapter.Remove(0, 8);
                }
            }*/
            #endregion

            #region Replace column C - CheckDegradationMode
            //rangeTo = rangeTo.Replace('H', 'C');
            rangeTo = "C1:C" + sheet.UsedRange.Rows.Count;
            var rowsC3 = GetRowsIndexesByText(sheet.Range["C1"], sheet.Range[rangeTo], "CheckDegradationMode");
            if (rowsC3 != null)
            {
                if (rowsC3.Count > 0)
                {
                    rangeB_To = "B1:B" + rowsC3[0];
                }

                var rowIncrement = 0;
                foreach (var row in rowsC3)
                {
                    // increment row (if new lines were added)
                    var index = row + rowIncrement;

                    if (exigence.DegradationMode.Count == 1 && exigence.DegradationMode[0] == ConfigurationManager.AppSettings["InputFile1_NoDegradationMode"])
                    {
                        Excel.Range rng = sheet.Range["C" + index];
                        rng.EntireRow.Delete(Type.Missing);
                        rowIncrement--;
                        continue;
                    }

                    rowIncrement = exigence.DegradationMode.Count - 1;
                    // insert 'n-1' new degradation modes lines
                    for (int i = 0; i < rowIncrement; i++)
                    {
                        sheet.Rows[index + 1].Insert();
                    }

                    // copy data to new lines
                    Excel.Range from = sheet.Range["C" + index].EntireRow;
                    Excel.Range to = sheet.Range["C" + index + ":C" + (index + rowIncrement)].EntireRow;
                    from.Copy(to);

                    for (int j = 0; j < rowIncrement + 1; j++)
                    {
                        sheet.Cells[index + j, 6] = exigence.DegradationMode[j];
                    }
                }
            }
            #endregion

            #region Replace column B - SetUp initial conditions, Check activation of diag before failure, Steps to create & remove the failure

            #region Replace 'SetUp initial conditions' - sheet InitialConditions
            rangeTo = "B1:B" + sheet.UsedRange.Rows.Count;
            var rowsC2 = GetRowsIndexesByText(sheet.Range["B1"], sheet.Range[rangeTo], "SetUp initial conditions");
            if (rowsC2 != null)
            {
                var rowIncrement = 0;
                foreach (var row in rowsC2)
                {
                    // increment row (if new lines were added)
                    var index = row + rowIncrement;

                    // number of new rows to be inserted
                    rowIncrement = failure.FailureDetails_InitialConditions.Count - 1;

                    // insert 'n-1' new failures procs lines
                    for (int i = 0; i < rowIncrement; i++)
                    {
                        sheet.Rows[index + 1].Insert();
                    }

                    // copy data to new lines
                    Excel.Range from = sheet.Range["B" + index].EntireRow;
                    Excel.Range to = sheet.Range["B" + index + ":B" + (index + rowIncrement)].EntireRow;
                    from.Copy(to);

                    for (int j = 0; j < rowIncrement + 1; j++)
                    {
                        sheet.Cells[index + j, 2] = failure.FailureDetails_InitialConditions[j].ToRealize;
                        sheet.Cells[index + j, 3] = failure.FailureDetails_InitialConditions[j].Step;
                        sheet.Cells[index + j, 4] = failure.FailureDetails_InitialConditions[j].CANFrame;
                        sheet.Cells[index + j, 5] = failure.FailureDetails_InitialConditions[j].CANMessage;
                        sheet.Cells[index + j, 6] = failure.FailureDetails_InitialConditions[j].ValueToBeGiven;
                        sheet.Cells[index + j, 7] = failure.FailureDetails_InitialConditions[j].VariableToCheck;
                        sheet.Cells[index + j, 8] = failure.FailureDetails_InitialConditions[j].ValueToBeChecked;
                    }
                }
            }
            #endregion

            // replace Check activation of diag before failure - sheet DiagEna
            rangeTo = "B1:B" + sheet.UsedRange.Rows.Count;

            rowsC2 = GetRowsIndexesByText(sheet.Range["B1"], sheet.Range[rangeTo], "Check activation of diag before failure");
            if (rowsC2 != null)
            {
                var rowIncrement = 0;
                foreach (var row in rowsC2)
                {
                    // increment row (if new lines were added)
                    var index = row + rowIncrement;

                    // number of new rows to be inserted
                    rowIncrement = failure.FailureDetails_DiagEna.Count - 1;

                    // insert 'n-1' new failures procs lines
                    for (int i = 0; i < rowIncrement; i++)
                    {
                        sheet.Rows[index + 1].Insert();
                    }

                    // copy data to new lines
                    Excel.Range from = sheet.Range["B" + index].EntireRow;
                    Excel.Range to = sheet.Range["B" + index + ":B" + (index + rowIncrement)].EntireRow;
                    from.Copy(to);

                    for (int j = 0; j < rowIncrement + 1; j++)
                    {
                        sheet.Cells[index + j, 2] = failure.FailureDetails_DiagEna[j].ToRealize;
                        sheet.Cells[index + j, 3] = failure.FailureDetails_DiagEna[j].Step;
                        sheet.Cells[index + j, 4] = failure.FailureDetails_DiagEna[j].CANFrame;
                        sheet.Cells[index + j, 5] = failure.FailureDetails_DiagEna[j].CANMessage;
                        sheet.Cells[index + j, 6] = failure.FailureDetails_DiagEna[j].ValueToBeGiven;
                        sheet.Cells[index + j, 7] = failure.FailureDetails_DiagEna[j].VariableToCheck;
                        sheet.Cells[index + j, 8] = failure.FailureDetails_DiagEna[j].ValueToBeChecked;
                    }
                }
            }

            rangeTo = "B1:B" + sheet.UsedRange.Rows.Count;
            rowsC2 = GetRowsIndexesByText(sheet.Range["B1"], sheet.Range[rangeTo], "Steps to create the failure");
            if (rowsC2 != null)
            {
                var rowIncrement = 0;
                foreach (var row in rowsC2)
                {
                    // increment row (if new lines were added)
                    var index = row + rowIncrement;

                    // number of new rows to be inserted
                    rowIncrement = failure.FailureDetails_Activate.Count - 1;
                    
                    // insert 'n-1' new failures procs lines
                    for (int i = 0; i < rowIncrement; i++)
                    {
                        sheet.Rows[index + 1].Insert();
                    }

                    // copy data to new lines
                    Excel.Range from = sheet.Range["B" + index].EntireRow;
                    Excel.Range to = sheet.Range["B" + index + ":B" + (index + rowIncrement)].EntireRow;
                    from.Copy(to);

                    for (int j = 0; j < rowIncrement + 1; j++)
                    {
                        sheet.Cells[index + j, 2] = failure.FailureDetails_Activate[j].ToRealize;
                        sheet.Cells[index + j, 3] = failure.FailureDetails_Activate[j].Step;
                        sheet.Cells[index + j, 4] = failure.FailureDetails_Activate[j].CANFrame;
                        sheet.Cells[index + j, 5] = failure.FailureDetails_Activate[j].CANMessage;
                        sheet.Cells[index + j, 6] = failure.FailureDetails_Activate[j].ValueToBeGiven;
                        sheet.Cells[index + j, 7] = failure.FailureDetails_Activate[j].VariableToCheck;
                        sheet.Cells[index + j, 8] = failure.FailureDetails_Activate[j].ValueToBeChecked;
                    }
                }
            }

            // replace steps to remove the failure - sheet deactivate
            rangeTo = "B1:B" + sheet.UsedRange.Rows.Count;
            rowsC2 = GetRowsIndexesByText(sheet.Range["B1"], sheet.Range[rangeTo], "Steps to remove the failure");
            if (rowsC2 != null)
            {
                var rowIncrement = 0;
                foreach (var row in rowsC2)
                {
                    // increment row (if new lines were added)
                    var index = row + rowIncrement;

                    // number of new rows to be inserted
                    rowIncrement = failure.FailureDetails_Deactivate.Count - 1;

                    // insert 'n-1' new failures procs lines
                    for (int i = 0; i < rowIncrement; i++)
                    {
                        sheet.Rows[index + 1].Insert();
                    }

                    // copy data to new lines
                    Excel.Range from = sheet.Range["B" + index].EntireRow;
                    Excel.Range to = sheet.Range["B" + index + ":B" + (index + rowIncrement)].EntireRow;
                    from.Copy(to);

                    for (int j = 0; j < rowIncrement + 1; j++)
                    {
                        sheet.Cells[index + j, 2] = failure.FailureDetails_Deactivate[j].ToRealize;
                        sheet.Cells[index + j, 3] = failure.FailureDetails_Deactivate[j].Step;
                        sheet.Cells[index + j, 4] = failure.FailureDetails_Deactivate[j].CANFrame;
                        sheet.Cells[index + j, 5] = failure.FailureDetails_Deactivate[j].CANMessage;
                        sheet.Cells[index + j, 6] = failure.FailureDetails_Deactivate[j].ValueToBeGiven;
                        sheet.Cells[index + j, 7] = failure.FailureDetails_Deactivate[j].VariableToCheck;
                        sheet.Cells[index + j, 8] = failure.FailureDetails_Deactivate[j].ValueToBeChecked;
                    }
                }
            }

            // if tdh null - remove 'Check increase of counter of driving cycles' row
            /*rangeTo = "B1:B" + sheet.UsedRange.Rows.Count;
            rowsC2 = GetRowsIndexesByText(sheet.Range["B1"], sheet.Range[rangeTo], "Check increase of counter of driving cycles");
            if (rowsC2 != null)
            {
                var rowIncrement = 0;
                foreach (var row in rowsC2)
                {
                    // modify row (if new lines were added/removed)
                    var index = row + rowIncrement;

                    if (homologation == null || homologation.FailingDiagCode == null)   // remove homologation line
                    {
                        Excel.Range rng = sheet.Range["B" + index];
                        rng.EntireRow.Delete(Type.Missing);
                        rowIncrement--;
                        continue;
                    }
                }
            }*/
            #endregion

            // merge cells in column A
            rangeA.Merge();
            rangeA.Interior.Color = indexRecord % 2 == 0 ? ColorTranslator.ToOle(Color.FromArgb(151, 157, 163)) : ColorTranslator.ToOle(Color.FromArgb(227, 159, 222));

            #region merge cells in column B
            try
            {
                var rowsNb = sheet.UsedRange.Rows.Count;
                var rangeStart = 1;
                var rangeEnd = 1;
                var lastValue = string.Empty;
                var newValue = string.Empty;

                for (int i = 1; i <= rowsNb; i++)
                {
                    newValue = Convert.ToString(sheet.Cells[i, 2].Value);
                    if (lastValue == newValue && i != rowsNb)
                    {
                        rangeEnd = i;
                    }
                    else
                    {
                        // if last line
                        if (i == rowsNb && lastValue == newValue)
                            rangeEnd = i;

                        if (rangeEnd > rangeStart)
                        {
                            var range = "B" + rangeStart + ":B" + rangeEnd;
                            sheet.Range[range].Merge();
                        }
                        lastValue = newValue;
                        rangeStart = i;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Merge on column B failed");
            }
            #endregion

            return sheet;
        }

        public static List<int> GetRowsIndexesByText(Excel.Range startRange, Excel.Range searchRange, string search)
        {
            List<int> matches = new List<int>();

            int i = 0;
            try
            {
                searchRange.Find(What: search, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns);

                Excel.Range next = startRange;

                //Logger.Log("searchRange count" + searchRange.Count);
                while (i++ < searchRange.Count)
                {
                    //next = searchRange.FindNext(next.Offset[1, 0]);
                    next = searchRange.FindNext(next);
                    if (next != null && !matches.Contains(next.Row))
                        matches.Add(next.Row);

                    //if (!matches.Add(next.Row))
                    //break;
                }
            }
            catch (Exception ex)
            {
                var sRange = startRange.Columns.Count + "," + startRange.Rows.Count;
                var eRange = searchRange.Columns.Count + "," + searchRange.Rows.Count;
                Logger.Log(string.Format("Error in GetRowsIndexByText ({0}) [{1}] : [{2}], i = {3} >> {4}", search, sRange, eRange, i, ex.Message + " - " + ex.StackTrace));
            }

            return matches;
        }

        #endregion
    }
}
