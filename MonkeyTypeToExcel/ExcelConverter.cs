using System.Windows.Forms;

namespace MonkeyTypeToExcel
{
    public partial class ExcelConverter : Form
    {
        public ExcelConverter()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new())
            {
                dialog.InitialDirectory = "c:\\";
                dialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    textBox1.Text = dialog.FileName;

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dialog = new())
            {
                dialog.InitialDirectory = "c:\\";
                dialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                dialog.FileName = "myExcelFile.csv";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    textBox2.Text = dialog.FileName;
                }
            }
        }

        public void UpdateProgressBar(int step, int length)
        {
            progressBar1.Value = (step/length)*100;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int step = 0;
            progressBar1.Value = 0;
            string csv = File.ReadAllText(textBox1.Text);
            string[] data = csv.Split("\n");
            step++; UpdateProgressBar(step, data.Length);

            string builtFile = "";
            for (int i = 1; i < data.Length; i++) 
            {
                string run = data[i];

                string[] values = run.Split("|");
                try
                {
                    long unixTimeMS = long.Parse(values[23]);
                    DateTimeOffset time = DateTimeOffset.FromUnixTimeMilliseconds(unixTimeMS);
                    string date = $"{time.Month}/{time.Day}/{time.Year}";
                    string addedLine = $"{values[2]},{values[4]},{values[3]}%,{values[5]}%,{char.ToUpper(values[16][0]) + values[16].Replace("_", " ").Substring(1)},{char.ToUpper(values[7][0]) + values[7].Substring(1)} {values[8]},{char.ToUpper(values[17][0]) + values[17].Substring(1)},{date}";
                    builtFile += addedLine + "\n";
                }
                catch (Exception) { }
                step++; UpdateProgressBar(step, data.Length);
            }
            File.WriteAllText(textBox2.Text, builtFile);
        }
    }
}