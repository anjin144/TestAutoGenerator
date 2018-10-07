using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TestAutoGenerator
{
    public static class AppManager
    {
        public static Hashtable FailingDiagCodes_Mappings = new Hashtable();

        static AppManager()
        {
            var mappings = ConfigurationManager.AppSettings["InputFile3_Column_FailingDiagCodes_Mappings"];
            if (!string.IsNullOrEmpty(mappings) && mappings.Contains(","))
            {
                var allValues = mappings.Split(',');
                foreach (var val in allValues)
                {
                    var key = val.Substring(0, 1);
                    var value = val.Substring(1);
                    FailingDiagCodes_Mappings.Add(key, value);
                }
            }
        }

        public static List<ExigenceDetails> GetExigences(string File1_Path, string sheetName, int exigencesColIndex)
        {
            var results = new List<ExigenceDetails>();

            try
            {
                #region read input file columns indexes

                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(File1_Path, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                var sheet = hssfwb.GetSheet(ConfigurationManager.AppSettings["InputFile1_SheetName"]);
                var rowHeaderIndex = int.Parse(ConfigurationManager.AppSettings["InputFile1_Header_RowIndex"]);

                var colFailureNameIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                    ConfigurationManager.AppSettings["InputFile1_Column_FailureName"]);
                var colDegradationModeIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                    ConfigurationManager.AppSettings["InputFile1_Column_DegradationMode"]);
                var colGIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                    ConfigurationManager.AppSettings["InputFile1_Column_G"]);
                var colGEEIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                    ConfigurationManager.AppSettings["InputFile1_Column_GEE"]);

                // check if columns FailureNameIndex and GIndex exists and return null if they're missing. 
                // Obs: column GEEIndex is not mandatory
                if (colFailureNameIndex == -1 || colGIndex == -1) 
                {
                    Logger.Log(string.Format("Input file '{0}' column(s) not found. Check the column names in config file!", File1_Path));
                    return null;
                }

                #endregion

                var colValueForSearch = "x";
                // get exicences from File1_Input
                var rowsDiags = ExcelUtils.GetRowsByColumnIndexAndCellValue(File1_Path, sheetName, exigencesColIndex, colValueForSearch);
                foreach (var rowDiag in rowsDiags)
                {
                    var exigence = new ExigenceDetails();
                    var isNewExicence = true;
                    try
                    {
                        exigence.FailureName = rowDiag.Cells[colFailureNameIndex].StringCellValue;

                        // check if failure already exists [update 4 feb 2018]
                        if (results.FirstOrDefault(ex => ex.FailureName == exigence.FailureName) != null)
                        {
                            //continue;
                            exigence = results.FirstOrDefault(ex => ex.FailureName == exigence.FailureName);
                            isNewExicence = false;
                        }

                        var degModes = rowDiag.Cells[colDegradationModeIndex].StringCellValue;
                        if (degModes != null)
                        {
                            if (degModes.Contains('\n'))
                            {
                                var newDegModes = degModes.Split('\n');
                                foreach (var degMode in newDegModes)
                                {
                                    if (exigence.DegradationMode.FirstOrDefault(d => d == degMode) != null)
                                        continue;

                                    if (degMode != ConfigurationManager.AppSettings["InputFile1_NoDegradationMode"])
                                        exigence.DegradationMode.Add(degMode);
                                }
                                //exigence.DegradationMode.AddRange(degModes.Split('\n'));
                            }
                            else
                            {
                                if (exigence.DegradationMode.FirstOrDefault(d => d == degModes) == null &&
                                    degModes != ConfigurationManager.AppSettings["InputFile1_NoDegradationMode"])
                                    exigence.DegradationMode.Add(degModes);
                            }
                        }

                        var voyant = rowDiag.Cells[colGIndex].StringCellValue;
                        if (!string.IsNullOrEmpty(voyant))
                        {
                            exigence.G1 = (voyant == "G1");
                            exigence.G2 = (voyant == "G2"); //rowDiag.Cells[colG2Index].StringCellValue;
                            //exigence.GEE = (voyant == "GEE"); //rowDiag.Cells[colGEEIndex].StringCellValue;
                        }

                        if (colGEEIndex != -1)
                        {
                            var gee = rowDiag.Cells[colGEEIndex].StringCellValue;
                            exigence.GEE = gee.ToUpper().Contains("BATTWARNREQ");
                        }

                        if (isNewExicence)
                            results.Add(exigence);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex, string.Format("Error extracting exigence details from file {0} at line {1}", File1_Path, rowDiag.RowNum));
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

            return results;
        }

        public static Failure GetFailureByName(string name, string sheetActivate, string sheetDeactivate, string sheetDiagEna, string sheetInitialConditions)
        {
            Failure result = null;
            try
            {
                result = new Failure(name);
                var filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + ConfigurationManager.AppSettings["InputFile2_Name"];
                var rowsFailureDetails = ExcelUtils.GetRowsByColumnIndexAndCellValue(filePath, sheetActivate, 0, name);
                var rowsFailureDetails_Deactivate = ExcelUtils.GetRowsByColumnIndexAndCellValue(filePath, sheetDeactivate, 0, name);
                var rowsFailureDetails_DiagEna = ExcelUtils.GetRowsByColumnIndexAndCellValue(filePath, sheetDiagEna, 0, name);
                var rowsFailureDetails_InitialConditions = ExcelUtils.GetRowsByColumnIndexAndCellValue(filePath, sheetInitialConditions, 0, name);

                #region rowsFailureDetails processing
                if (rowsFailureDetails != null && rowsFailureDetails.Count > 0)
                {
                    foreach (var row in rowsFailureDetails)
                    {
                        var detail = new FailureDetail();
                        detail.ToRealize = row.Cells[1].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[1].StringCellValue :
                                           row.Cells[1].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[1].NumericCellValue.ToString() : "";

                        detail.Step = row.Cells[2].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[2].StringCellValue :
                                      row.Cells[2].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[2].NumericCellValue.ToString() : "";

                        detail.CANFrame = row.Cells[3].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[3].StringCellValue :
                                          row.Cells[3].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[3].NumericCellValue.ToString() : "";

                        detail.CANMessage = row.Cells[4].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[4].StringCellValue :
                                            row.Cells[4].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[4].NumericCellValue.ToString() : "";

                        detail.ValueToBeGiven = row.Cells[5].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[5].StringCellValue :
                                                row.Cells[5].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[5].NumericCellValue.ToString() : "";

                        detail.VariableToCheck = row.Cells[6].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[6].StringCellValue :
                                                 row.Cells[6].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[6].NumericCellValue.ToString() : "";

                        detail.ValueToBeChecked = row.Cells[7].CellType == NPOI.SS.UserModel.CellType.String ? row.Cells[7].StringCellValue :
                                                  row.Cells[7].CellType == NPOI.SS.UserModel.CellType.Numeric ? row.Cells[7].NumericCellValue.ToString() : "";
                        
                        //add detail to parent
                        result.FailureDetails_Activate.Add(detail);
                    }
                }
                #endregion

                #region rowsFailureDetails_Deactivate processing
                if (rowsFailureDetails_Deactivate != null && rowsFailureDetails_Deactivate.Count > 0)
                {
                    foreach (var row in rowsFailureDetails_Deactivate)
                    {
                        var detail = new FailureDetail();
                        detail.ToRealize = row.Cells[1].CellType == CellType.String ? row.Cells[1].StringCellValue :
                                           row.Cells[1].CellType == CellType.Numeric ? row.Cells[1].NumericCellValue.ToString() : "";

                        detail.Step = row.Cells[2].CellType == CellType.String ? row.Cells[2].StringCellValue :
                                      row.Cells[2].CellType == CellType.Numeric ? row.Cells[2].NumericCellValue.ToString() : "";

                        detail.CANFrame = row.Cells[3].CellType == CellType.String ? row.Cells[3].StringCellValue :
                                          row.Cells[3].CellType == CellType.Numeric ? row.Cells[3].NumericCellValue.ToString() : "";

                        detail.CANMessage = row.Cells[4].CellType == CellType.String ? row.Cells[4].StringCellValue :
                                            row.Cells[4].CellType == CellType.Numeric ? row.Cells[4].NumericCellValue.ToString() : "";

                        detail.ValueToBeGiven = row.Cells[5].CellType == CellType.String ? row.Cells[5].StringCellValue :
                                                row.Cells[5].CellType == CellType.Numeric ? row.Cells[5].NumericCellValue.ToString() : "";

                        detail.VariableToCheck = row.Cells[6].CellType == CellType.String ? row.Cells[6].StringCellValue :
                                                 row.Cells[6].CellType == CellType.Numeric ? row.Cells[6].NumericCellValue.ToString() : "";

                        detail.ValueToBeChecked = row.Cells[7].CellType == CellType.String ? row.Cells[7].StringCellValue :
                                                  row.Cells[7].CellType == CellType.Numeric ? row.Cells[7].NumericCellValue.ToString() : "";

                        //add detail to parent
                        result.FailureDetails_Deactivate.Add(detail);
                    }
                }
                #endregion

                #region rowsFailureDetails_DiagEna processing
                if (rowsFailureDetails_DiagEna != null && rowsFailureDetails_DiagEna.Count > 0)
                {
                    foreach (var row in rowsFailureDetails_DiagEna)
                    {
                        var detail = new FailureDetail();
                        detail.ToRealize = row.Cells[1].CellType == CellType.String ? row.Cells[1].StringCellValue :
                                           row.Cells[1].CellType == CellType.Numeric ? row.Cells[1].NumericCellValue.ToString() : "";

                        detail.Step = row.Cells[2].CellType == CellType.String ? row.Cells[2].StringCellValue :
                                      row.Cells[2].CellType == CellType.Numeric ? row.Cells[2].NumericCellValue.ToString() : "";

                        detail.CANFrame = row.Cells[3].CellType == CellType.String ? row.Cells[3].StringCellValue :
                                          row.Cells[3].CellType == CellType.Numeric ? row.Cells[3].NumericCellValue.ToString() : "";

                        detail.CANMessage = row.Cells[4].CellType == CellType.String ? row.Cells[4].StringCellValue :
                                            row.Cells[4].CellType == CellType.Numeric ? row.Cells[4].NumericCellValue.ToString() : "";

                        detail.ValueToBeGiven = row.Cells[5].CellType == CellType.String ? row.Cells[5].StringCellValue :
                                                row.Cells[5].CellType == CellType.Numeric ? row.Cells[5].NumericCellValue.ToString() : "";

                        detail.VariableToCheck = row.Cells[6].CellType == CellType.String ? row.Cells[6].StringCellValue :
                                                 row.Cells[6].CellType == CellType.Numeric ? row.Cells[6].NumericCellValue.ToString() : "";

                        detail.ValueToBeChecked = row.Cells[7].CellType == CellType.String ? row.Cells[7].StringCellValue :
                                                  row.Cells[7].CellType == CellType.Numeric ? row.Cells[7].NumericCellValue.ToString() :
                                                  row.Cells[7].CellType == CellType.Boolean ? row.Cells[7].BooleanCellValue.ToString() : "";

                        //add detail to parent
                        result.FailureDetails_DiagEna.Add(detail);
                    }
                }
                #endregion

                #region rowsFailureDetails_InitialConditions processing
                if (rowsFailureDetails_InitialConditions != null && rowsFailureDetails_InitialConditions.Count > 0)
                {
                    foreach (var row in rowsFailureDetails_InitialConditions)
                    {
                        var detail = new FailureDetail();
                        detail.ToRealize = row.Cells[1].CellType == CellType.String ? row.Cells[1].StringCellValue :
                                           row.Cells[1].CellType == CellType.Numeric ? row.Cells[1].NumericCellValue.ToString() : "";

                        detail.Step = row.Cells[2].CellType == CellType.String ? row.Cells[2].StringCellValue :
                                      row.Cells[2].CellType == CellType.Numeric ? row.Cells[2].NumericCellValue.ToString() : "";

                        detail.CANFrame = row.Cells[3].CellType == CellType.String ? row.Cells[3].StringCellValue :
                                          row.Cells[3].CellType == CellType.Numeric ? row.Cells[3].NumericCellValue.ToString() : "";

                        detail.CANMessage = row.Cells[4].CellType == CellType.String ? row.Cells[4].StringCellValue :
                                            row.Cells[4].CellType == CellType.Numeric ? row.Cells[4].NumericCellValue.ToString() : "";

                        detail.ValueToBeGiven = row.Cells[5].CellType == CellType.String ? row.Cells[5].StringCellValue :
                                                row.Cells[5].CellType == CellType.Numeric ? row.Cells[5].NumericCellValue.ToString() : "";

                        detail.VariableToCheck = row.Cells[6].CellType == CellType.String ? row.Cells[6].StringCellValue :
                                                 row.Cells[6].CellType == CellType.Numeric ? row.Cells[6].NumericCellValue.ToString() : "";

                        detail.ValueToBeChecked = row.Cells[7].CellType == CellType.String ? row.Cells[7].StringCellValue :
                                                  row.Cells[7].CellType == CellType.Numeric ? row.Cells[7].NumericCellValue.ToString() :
                                                  row.Cells[7].CellType == CellType.Boolean ? row.Cells[7].BooleanCellValue.ToString() : "";

                        //add detail to parent
                        result.FailureDetails_InitialConditions.Add(detail);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                Logger.Log(ex, string.Format("Error reading failure '{0}' details from file '{1}'.", name, ConfigurationManager.AppSettings["InputFile2_Name"]));
            }

            return result;
        }

        public static Homologation GetHomologation(string filePath, string sheetName, string failureName)
        {
            Homologation result = null;

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            var sheet = hssfwb.GetSheet(sheetName);

            var rowHeaderIndex = int.Parse(ConfigurationManager.AppSettings["InputFile3_Header_RowIndex"]);

            var colFailureNameIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                ConfigurationManager.AppSettings["InputFile3_Column_Exigences"]);
            var colMILndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                ConfigurationManager.AppSettings["InputFile3_Column_MIL"]);
            var colFailingDiagCodeIndex = ExcelUtils.GetColumnIndexByText(sheet, rowHeaderIndex,
                                                                ConfigurationManager.AppSettings["InputFile3_Column_DrivingCycles"]);

            if (colFailureNameIndex == -1 || colMILndex == -1 || colFailingDiagCodeIndex == -1)
            {
                Logger.Log(string.Format("Input file '{0}' column(s) not found. Check the column names in config file!", filePath));
                return null;
            }

            var rowsCol0 = ExcelUtils.GetRowsByColumnIndexAndCellValue(sheet, colFailureNameIndex, failureName);
            if (rowsCol0 != null && rowsCol0.Count > 0 && rowsCol0[0] != null)
            {
                var data = rowsCol0[0];

                result = new Homologation();
                result.FailureName = failureName;
                result.FailingDiagCode = data.Cells[colFailingDiagCodeIndex].CellType == CellType.String ? data.Cells[colFailingDiagCodeIndex].StringCellValue :
                                         data.Cells[colFailingDiagCodeIndex].CellType == CellType.Numeric ? data.Cells[colFailingDiagCodeIndex].NumericCellValue.ToString() : "";
                result.MIL = data.Cells[colMILndex].CellType == CellType.String ? data.Cells[colMILndex].StringCellValue :
                             data.Cells[colMILndex].CellType == CellType.Numeric ? data.Cells[colMILndex].NumericCellValue.ToString() : "";
            }
            return result;
        }

        public static void ProcessFailure(Microsoft.Office.Interop.Excel.Application excelApplication, string templateFilePath, string outputFilePath, ExigenceDetails exigence, Failure failure, Homologation homologation, int indexRecordProcessed)
        {
            ExcelUtils.WriteFailureDataToExcel(excelApplication, templateFilePath, outputFilePath, exigence, failure, homologation, indexRecordProcessed);
        }
    }
}
