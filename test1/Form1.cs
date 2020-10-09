using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace test1
{
    public partial class Form1 : Form
    {
        private Pke _pke;

        public Form1()
        {
            InitializeComponent();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string PathFolder = string.Empty;

            while (true)
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

                folderBrowserDialog.Description = @"Укажите директорию PKE.";

                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    PathFolder = folderBrowserDialog.SelectedPath;

                    if (Path.GetFileName(PathFolder) == "PKE") break;

                    MessageBox.Show(@"Указанная директория не являектся PKE.", @"Ошибка.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    return;
                }
            }

            _pke = new Pke(PathFolder);

            if (_pke.MessageException != string.Empty)
            {
                MessageBox.Show($@"Возникли следующие ошибки:
{_pke.MessageException}", @"Ошибка.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            dataGridView1.DataSource = _pke.GetParamTable;
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            while (dataGridView1.Rows.Count > 1)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    dataGridView1.Rows.Remove(row);
                }
            }

            dataGridView2.DataSource = null;
        }

        private void выгрузитьВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_pke == null)
            {
                MessageBox.Show(@"Выберите измерение.", @"Предупреждение.", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = @"Text files(*.xlsx)|*.xlsx|All files(*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.Cancel) return;

            string excelFile = saveFileDialog.FileName;

            _pke.ExportToExcel(excelFile);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Pke.CreateParamTable();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = ((DataGridView)sender).SelectedCells[0].RowIndex;

            if (_pke != null && rowIndex < _pke.Length) dataGridView2.DataSource = _pke.GetResultTable(rowIndex);
        }
    }
}
