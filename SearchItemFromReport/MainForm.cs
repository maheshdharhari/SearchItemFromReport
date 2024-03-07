using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace SearchFromReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            cmbItemType.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = @"Browse Report File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "html",
                Filter = @"Html Files (*.html)|*.html",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            _itemToSearch = cmbItemType.SelectedItem.ToString();
            var htmlFilePath = textBox1.Text;
            if (!File.Exists(htmlFilePath))
                return;
            var failedItems = await FindIdInHtml(htmlFilePath);
            var newFileWithExtension = _itemToSearch + ".csv";
            if (File.Exists(newFileWithExtension))
                File.Delete(newFileWithExtension);
            var stringValue = string.Join(Environment.NewLine, failedItems.ToArray());

            var saveFileDialog = new SaveFileDialog
            {
                InitialDirectory = @"D:\",
                Title = @"Browse Report File",

                CheckPathExists = true,

                DefaultExt = "html",
                Filter = @"CSV Files (*.csv)|*.csv",
                FilterIndex = 2,
                RestoreDirectory = true,
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                newFileWithExtension = saveFileDialog.FileName;
            }

            using (var sw = new StreamWriter(newFileWithExtension))
            {
                await sw.WriteAsync(stringValue);
            }

            MessageBox.Show(@"Created csv file for the items.", this.Name, MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            button3.Enabled = true;
        }

        private string _itemToSearch;

        private async Task<List<string>> FindIdInHtml(string htmlFilePath)
        {
            var allListOfItemId = new List<string>();
            await Task.Run(() =>
            {
                if (!File.Exists(htmlFilePath))
                    return null;
                var doc = new HtmlDocument();
                doc.Load(htmlFilePath);
                var tableRows = doc.DocumentNode.SelectNodes("//table[@id='myTable']//tr");
                foreach (var row in tableRows)
                {
                    try
                    {
                        var itemType = row.SelectNodes("td")[3].InnerText;
                        if (_itemToSearch != "All" && itemType != _itemToSearch) continue;
                        var itemName = row.SelectNodes("td")[0].InnerText;
                        var itemPath = row.SelectNodes("td")[1].InnerText;
                        var itemId = row.SelectNodes("td")[2].InnerText;
                        var itemWarningError = row.SelectNodes("td")[7].InnerText;
                        allListOfItemId.Add(itemId + "," + itemName + "," + itemPath + "," + itemWarningError);
                    }
                    catch (Exception)
                    {
                        // throw;
                    }
                }

                var unused = allListOfItemId.Count();

                return allListOfItemId;
            });
            return allListOfItemId;
        }
    }
}