using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace TestAutoGenerator
{
    public partial class Form1 : Form
    {
        // global var
        string FileSdF_Path = null;
        string FileHomologation_Path = null;
        string File1_Path_Output = null;

        public Form1()
        {
            InitializeComponent();

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
        }

        #region GUI events

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter options and filter index.
            dlg.Filter = "Text Files (.xls)|*.xlsx|All Files (*.*)|*.*";
            dlg.FilterIndex = 1;

            dlg.Multiselect = true;

            // Call the ShowDialog method to show the dialog box.
            var userClickedOK = dlg.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                FileSdF_Path = dlg.FileName;
                txtFile1.Text = FileSdF_Path.Substring(FileSdF_Path.LastIndexOf('\\') + 1);

                var shortfileName = txtFile1.Text.Substring(0, txtFile1.Text.LastIndexOf("."));
                File1_Path_Output = FileSdF_Path.Replace(shortfileName, shortfileName + "_out");

                var exigencesColumn = -1;
                int.TryParse(ConfigurationManager.AppSettings["InputFile1_Column_Exigences"], out exigencesColumn); //ExcelUtils.GetColumnIndexByText(FileSdF_Path, ConfigurationManager.AppSettings["InputFile1_Column_Exigences"]);
                if (exigencesColumn == -1)
                {
                    MessageBox.Show("Incorrect starting index set for applicability columns!");
                    return;
                }

                var nextColumns = ExcelUtils.GetColumnsAfterIndex(FileSdF_Path, exigencesColumn, int.Parse(ConfigurationManager.AppSettings["InputFile1_Header_RowIndex"]));
                cbApplications.DataSource = new BindingSource(nextColumns, null);
                cbApplications.DisplayMember = "Value";
                cbApplications.ValueMember = "Key";
            }
        }

        private void btnSelectFile_Homologation_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter options and filter index.
            dlg.Filter = "Text Files (.xls)|*.xlsx|All Files (*.*)|*.*";
            dlg.FilterIndex = 1;

            dlg.Multiselect = true;

            // Call the ShowDialog method to show the dialog box.
            var userClickedOK = dlg.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                FileHomologation_Path = dlg.FileName;
                txtHomologation.Text = FileHomologation_Path.Substring(FileHomologation_Path.LastIndexOf('\\') + 1);
            }
        }
        private void chKHomologation_CheckedChanged(object sender, EventArgs e)
        {
            btnSelectFile_Homologation.Enabled = chKHomologation.Checked;
            if (!chKHomologation.Checked)
                txtHomologation.Text = string.Empty;
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            ConfigForm popup = new ConfigForm();
            DialogResult dialogresult = popup.ShowDialog();
            if (dialogresult == DialogResult.OK)
            {
                Console.WriteLine("You clicked OK");
            }
            else if (dialogresult == DialogResult.Cancel)
            {
                Console.WriteLine("You clicked either Cancel or X button in the top right corner");
            }
            popup.Dispose();
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (!ValidateInputData())
                return;

            EnableDisableButtons(false);

            // Start the BackgroundWorker.
            backgroundWorker1.RunWorkerAsync();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            EnableDisableButtons(true);
            ClearInputControls();

            progressBar1.Value = 0;
            lblProgress.Text = string.Empty;
            this.Text = "Diag automatic test generator";
            ReportProgress("Prepare for a new processing.........", true, false);

            // reinit logger
            ConstructorInfo constructor = typeof(Logger).GetConstructor(BindingFlags.Static | BindingFlags.NonPublic, null, new Type[0], null);
            constructor.Invoke(null, null);
        }

        #endregion

        #region BackgroundWorker methods
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ReportProgress("***************************************************************************", true, true);
            ReportProgress("Start processing file '" + FileSdF_Path, true, true);

            int exigencesColIndex = 0;// ((KeyValuePair<int, string>)cbApplications.SelectedItem).Key;
            string selectedExigenceText = ""; // ((KeyValuePair<int, string>)cbApplications.SelectedItem).Value;

            MethodInvoker mi = delegate {
                exigencesColIndex = ((KeyValuePair<int, string>)cbApplications.SelectedItem).Key;
                selectedExigenceText = ((KeyValuePair<int, string>)cbApplications.SelectedItem).Value;
            };
            if (this.InvokeRequired)
                this.Invoke(mi);

            ReportProgress("Selected exigence: " + exigencesColIndex + " - " + selectedExigenceText, true, true);
            ReportProgress("***************************************************************************", true, true);

            ReportProgress("Start extracting exigences from input file.", true, true);
            var sheetName = ConfigurationManager.AppSettings["InputFile1_SheetName"];
            var exigences = AppManager.GetExigences(FileSdF_Path, sheetName, exigencesColIndex);  // exicencesColIndex + nb of blank columns at the start of the file
            ReportProgress("End extracting exigences from input file.", true, true);
            ReportProgress("***************************************************************************", true, true);

            if (exigences == null)
            {
                ReportProgress("No exigences found in your input file. Please check application log for details.", true, false);
                MessageBox.Show("No exigences found in your input file. Please check application log for details.");
                return;
            }

            #region Create the output file based on the template
            var appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var templateFilePath = appPath + "\\" + ConfigurationManager.AppSettings["OutputFile_Template"];
            var outputFilePath = appPath + "\\" + "Output_" + Logger.ExecutionTimeString + ".xlsx";
            File.Copy(templateFilePath, outputFilePath);

            // remove template lines from output file (let only the header) & copy sheet with gasoline/diesel
            string sheetToDelete = rdGasoline.Checked ? ConfigurationManager.AppSettings["InputFile2_SheetName_MDS_Diesel"] : ConfigurationManager.AppSettings["InputFile2_SheetName_MDS_Gazoline"];
            ExcelUtils.InitiateOutputFile(outputFilePath, sheetToDelete);
            #endregion

            var noDegrdationModeExigences = new StringBuilder();
            var noProcedureAvailableExigences = new StringBuilder();

            var excelApplication = new Microsoft.Office.Interop.Excel.Application();
            excelApplication.Application.DisplayAlerts = false;

            int noDegradationModeCount = 0;
            int noProcedureAvailableCount = 0;
            int noExigencesProcessed = 0;
            int noItemProcessed = 0;

            #region Process exigences
            foreach (var exigence in exigences)
            {
                if (backgroundWorker1.CancellationPending == true)
                {
                    // close the excelApplication instance
                    ReportProgress("The process has been cancelled by the user!", true, true, true);
                    excelApplication.Quit();

                    DisplayProcessingSummary(exigences.Count, noDegradationModeCount, noProcedureAvailableCount, noExigencesProcessed);
                    Process.Start(outputFilePath);

                    // re-enable the button 'New'
                    mi = delegate { btnNew.Enabled = true; };
                    if (this.InvokeRequired)
                        this.Invoke(mi);

                    // cancel and exit worker
                    e.Cancel = true;
                    return;
                }

                noItemProcessed++;
                backgroundWorker1.ReportProgress(noItemProcessed * 100 / exigences.Count);

                if (exigence.DegradationMode.Count == 1 &&
                    exigence.DegradationMode[0] == ConfigurationManager.AppSettings["InputFile1_NoDegradationMode"])
                {
                    noDegradationModeCount++;
                    noDegrdationModeExigences.AppendLine(exigence.FailureName);
                    //continue;
                }

                var failure = AppManager.GetFailureByName(exigence.FailureName, 
                    ConfigurationManager.AppSettings["InputFile2_SheetName_Activate"],
                    ConfigurationManager.AppSettings["InputFile2_SheetName_Deactivate"], 
                    ConfigurationManager.AppSettings["InputFile2_SheetName_DiagEna"],
                    ConfigurationManager.AppSettings["InputFile2_SheetName_InitialConditions"]);

                if (failure.FailureDetails_Activate.Count == 0 ||
                   (failure.FailureDetails_Activate.Count == 1 && failure.FailureDetails_Activate[0].Step == ""))
                {
                    noProcedureAvailableCount++;
                    noProcedureAvailableExigences.AppendLine(exigence.FailureName);
                    continue;
                }

                ReportProgress(string.Format("Failure details: {0}, {1}", failure.Chapter, failure.FailureDetails_Activate.Count), false, true);

                #region read homologations
                Homologation homologation = null;
                if (chKHomologation.Checked)
                {
                    homologation = AppManager.GetHomologation(FileHomologation_Path, ConfigurationManager.AppSettings["InputFile3_SheetName"], failure.Chapter);
                }
                #endregion

                ReportProgress("Start processing failure " + failure.Chapter, true, false);
                AppManager.ProcessFailure(excelApplication, templateFilePath, outputFilePath, exigence, failure, homologation, noExigencesProcessed);
                noExigencesProcessed++;
                ReportProgress("End processing failure. Remaining items to be processed: " + (exigences.Count - noItemProcessed), true, false);
            }
            #endregion

            // finalise processing - add end line to output
            ExcelUtils.FinishOutputFile(outputFilePath);

            // close the excelApplication instance
            excelApplication.Quit();

            ReportProgress("End processing.", true, true);

            DisplayProcessingSummary(exigences.Count, noDegradationModeCount, noProcedureAvailableCount, noExigencesProcessed);

            if (noDegrdationModeExigences.Length > 0)
            {
                Logger.Log(string.Format("The following failures do not trigger any degraded mode: {0} {1}", Environment.NewLine, noDegrdationModeExigences.ToString()));
            }
            if (noProcedureAvailableExigences.Length > 0)
            {
                Logger.Log(string.Format("The following failures do not have any procedure available: {0} {1}", Environment.NewLine, noProcedureAvailableExigences.ToString()));
            }

            // open the output excel file
            Process.Start(outputFilePath);

            // re-enable the button 'New'
            mi = delegate {
                btnNew.Enabled = true;
                btnCancel.Enabled = false;
            };
            if (this.InvokeRequired)
                this.Invoke(mi);
        }
        
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar1.Value = e.ProgressPercentage;
            // Set the text.
            this.Text = "Diag automatic test generator: " + e.ProgressPercentage + "%";
            lblProgress.Text = e.ProgressPercentage + "%";
        }

        #endregion

        #region Private methods

        private bool ValidateInputData()
        {
            if (cbApplications.SelectedItem == null)
            {
                MessageBox.Show("Please select an application from SdF file before start the processing.");
                return false;
            }

            if (chKHomologation.Checked && string.IsNullOrEmpty(txtHomologation.Text))
            {
                MessageBox.Show("Please select Tdh file before start the processing.");
                return false;
            }

            if (!rdGasoline.Checked && !rdDiesel.Checked)
            {
                MessageBox.Show("Please select the fuel type before start the processing.");
                return false;
            }

            return true;
        }

        private void ReportProgress(string message, bool displayInProgress, bool log, bool isWarning = false)
        {
            if (displayInProgress)
            {
                MethodInvoker mi = delegate {
                    
                    if (isWarning)
                    {
                        txtProgress.SelectionFont = new Font("Verdana", 10, FontStyle.Bold);
                        txtProgress.SelectionColor = Color.Red;
                    }
                    txtProgress.AppendText(message + Environment.NewLine);
                };
                if (this.InvokeRequired)
                    this.Invoke(mi);
            }
            
            if (log)
                Logger.Log(message);
        }

        private void EnableDisableButtons(bool isEnabled)
        {
            btnConfig.Enabled = isEnabled;
            btnSelectFile.Enabled = isEnabled;
            chKHomologation.Enabled = isEnabled;
            btnSelectFile_Homologation.Enabled = false;
            rdGasoline.Enabled = isEnabled;
            rdDiesel.Enabled = isEnabled;
            btnProcess.Enabled = isEnabled;
            btnNew.Enabled = isEnabled;
            cbApplications.Enabled = isEnabled;

            if (isEnabled)
                btnCancel.Enabled = true;
        }

        private void ClearInputControls()
        {
            txtFile1.Text = string.Empty;
            cbApplications.DataSource = null;
            cbApplications.Text = "Select application";
            chKHomologation.Checked = false;
            txtHomologation.Text = string.Empty;
            rdGasoline.Checked = false;
            rdDiesel.Checked = false;
        }

        private void DisplayProcessingSummary(int totalNbOfItems, int nbDegradationModeCount, int nbProcedureAvailableCount, int nbExigencesProcessed)
        {
            ReportProgress("***************************************************************************", true, true);
            ReportProgress("Total number of exigences: " + totalNbOfItems, true, true);
            ReportProgress("Number of exigences which do not trigger any degraded mode: " + nbDegradationModeCount, true, true);
            ReportProgress("Number of exigences which do not have any procedure available: " + nbProcedureAvailableCount, true, true);
            ReportProgress("Number of exigences processed: " + nbExigencesProcessed, true, true);
            ReportProgress("***************************************************************************", true, true);
        }

        #endregion
    }
}
