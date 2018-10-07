using System.Configuration;
using System.Windows.Forms;

namespace TestAutoGenerator
{
    public partial class ConfigForm : Form
    {
        public ConfigForm()
        {
            InitializeComponent();

            LoadSettings();
        }

        private void LoadSettings()
        {            
            txtExigences_Sheet.Text = ConfigurationManager.AppSettings["InputFile1_SheetName"];
            txtExigences_hRowIndex.Text = ConfigurationManager.AppSettings["InputFile1_Header_RowIndex"];
            txtExigences_ColExigences.Text = ConfigurationManager.AppSettings["InputFile1_Column_Exigences"];
         }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ConfigurationManager.AppSettings.Set("InputFile1_SheetName", txtExigences_Sheet.Text);
            ConfigurationManager.AppSettings.Set("InputFile1_Header_RowIndex", txtExigences_hRowIndex.Text);
            ConfigurationManager.AppSettings.Set("InputFile1_Column_Exigences", txtExigences_ColExigences.Text);
        }
    }
}
