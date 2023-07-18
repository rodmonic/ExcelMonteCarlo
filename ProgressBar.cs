using ExcelDna.Integration;
using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace Campagna
{
    public partial class ProgressBar : Form
    {
        private bool _cancelled = false;

        public ProgressBar()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(ProgressBarClosing);
        }

        public void Update(int levelOnePercentage, int levelTwoPercentage, string levelOneText, string levelTwoText)
        {

            if ((bool)XlCall.Excel(XlCall.xlAbort, true) == true || _cancelled == true)
            {
                _cancelled = false;
                Globals._sliceData.Clear();
                throw new OperationCanceledException();
            }
            if (levelOnePercentage > 100 || levelOnePercentage < 0 || levelTwoPercentage > 100 || levelTwoPercentage < 0)
                throw new ArgumentOutOfRangeException("Progress can not be negative or > 100");

            levelOneLabel.Text = levelOneText;
            levelTwoLabel.Text = levelTwoText;
            progressBarLevelOne.Value = levelOnePercentage;
            progressBarLevelTwo.Value = levelTwoPercentage;
            Application.DoEvents();
            Debug.WriteLine($"{levelOneText}: {DateTime.Now - Globals._start}");

        }

        private void ProgressBarClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                _cancelled = true;
                Hide();

            }

        }

        private void ProgressBar_Load(object sender, EventArgs e)
        {

        }
    }
}
