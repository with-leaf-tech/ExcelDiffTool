using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDiffTool {
    public partial class ExcelDiffForm : Form {
        public ExcelDiffForm() {
            InitializeComponent();
            textBox3.Height = this.Height - 110;
            radioButton1.Checked = true;
        }

        private void Form1_Resize(object sender, EventArgs e) {
            textBox3.Height = this.Height - 110;
        }

        private void button1_Click(object sender, EventArgs e) {
            openFileDialog(textBox1);
        }

        private void button2_Click(object sender, EventArgs e) {
            openFileDialog(textBox2);
        }

        private void openFileDialog(TextBox textbox) {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //ofd.FileName = "default.html";
            //ofd.InitialDirectory = @"C:\";
            //ofd.Filter = "HTMLファイル(*.html;*.htm)|*.html;*.htm|すべてのファイル(*.*)|*.*";
            ofd.Filter = @"Excel ファイル|*.xls;*.xlsx;*.xlsm|全てのファイル|*.*";
            ofd.FilterIndex = 1;
            ofd.Title = "開くファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK) {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                textbox.Text = ofd.FileName;
            }
        }

        private async void button3_Click(object sender, EventArgs e) {
            textBox3.Text = "";
            await Task.Run(() => checkFiles(textBox1.Text, textBox2.Text));
        }

        private void checkFiles(string file1, string file2) {
            ParallelOptions options = new ParallelOptions();
            options.MaxDegreeOfParallelism = 4;


            StringBuilder sb = new StringBuilder();

            if (!File.Exists(file1)) {
                MessageBox.Show("ファイル1が存在しません");
                return;
            }
            if (!File.Exists(file2)) {
                MessageBox.Show("ファイル2が存在しません");
                return;
            }
            if (file1 == file2) {
                MessageBox.Show("ファイル1とファイル2が同一です");
                return;
            }
            updateStatus(false);




            Workbook workbook1 = new Workbook();
            Workbook workbook2 = new Workbook();

            workbook1.LoadFromFile(file1);
            workbook2.LoadFromFile(file2);

            if (workbook1.Worksheets.Count != workbook2.Worksheets.Count) {
                appendText("シート数が違います\tファイル1=" + workbook1.Worksheets.Count + "\tファイル2=" + workbook2.Worksheets.Count + Environment.NewLine);
                //sb.Append("シート数が違います\tファイル1=" + workbook1.Worksheets.Count + "\tファイル2=" + workbook2.Worksheets.Count + Environment.NewLine);
                //textBox3.Text = sb.ToString();
                updateStatus(true);
                return;
            }

            if (!radioButton1.Checked) {
                appendText("シート名\tセル\t計算式1\t計算式2" + Environment.NewLine);
            }
            else {
                appendText("シート名\tセル\t値1\t値2" + Environment.NewLine);
            }

            for (int i = 0; i < workbook1.Worksheets.Count; i++) {
                Worksheet sheet1 = workbook1.Worksheets[i];
                Worksheet sheet2 = workbook2.Worksheets[i];

                int maxRows = sheet1.Rows.Length;
                int maxColumns = sheet1.Columns.Length;

                if (sheet1.Columns.Length != sheet2.Columns.Length || sheet1.Rows.Length != sheet2.Rows.Length) {
                    appendText("レイアウトが違います\tシート=[" + sheet1.Name + "]\t" + sheet1.Columns.Length + "×" + sheet1.Rows.Length + "\t" + sheet1.Columns.Length + "×" + sheet1.Rows.Length + Environment.NewLine);
                    //sb.Append("レイアウトが違います シート=[" + sheet1.Name + "]\tファイル1=[" + sheet1.Columns.Length + "×" + sheet1.Rows.Length + "]\tファイル2=[" + sheet1.Columns.Length + "×" + sheet1.Rows.Length + "]" + Environment.NewLine);
                    //textBox3.Text = sb.ToString();
                    updateStatus(true);
                    return;
                }

                //CellRange aa = sheet1.Range[0, 0, maxRows, maxColumns];
                //aa.Cells[]

                for (int j = 1; j < maxRows+1; j++) {
                    this.Invoke((MethodInvoker)delegate {
                        label1.Text = "チェック中...シート[" + sheet1.Name + "]" + Environment.NewLine + "maxRows=" + maxRows + " Row =" + j;
                    });
                    //Parallel.For(1, maxColumns + 1, options, k => {
                    for (int k = 1; k < maxColumns+1; k++) {

                        CellRange cell1 = sheet1.Range[j, k];
                        CellRange cell2 = sheet2.Range[j, k];
                        if (!radioButton1.Checked) {
                            if (cell1.Formula != cell2.Formula) {
                                appendText(sheet1.Name + "\t" + cell1.RangeAddressLocal + "\t" + cell1.Formula + "\t" + cell2.Formula + Environment.NewLine);
                                //sb.Append("計算式が違います\tシート[" + sheet1.Name + "]" + "\tファイル1[" + j + "," + k + "]=" + cell1.Formula + "\tファイル2[" + j + "," + k + "]=" + cell2.Formula + Environment.NewLine);
                            }
                        }
                        else {
                            if (cell1.DisplayedText != cell2.DisplayedText) {
                                appendText(sheet1.Name + "\t" + cell1.RangeAddressLocal + "\t" + cell1.DisplayedText + "\t" + cell2.DisplayedText + Environment.NewLine);
                                //sb.Append("値が違います\tシート[" + sheet1.Name + "]" + "\tファイル1[" + j + "," + k + "]=" + cell1.Formula + "\tファイル2[" + j + "," + k + "]=" + cell2.Formula + Environment.NewLine);
                            }
                        }
                        }
                    //});
                }
            }

            updateStatus(true);
            this.Invoke((MethodInvoker)delegate {
                label1.Text = "";
            });
            MessageBox.Show("差分検出が終了しました");


            //textBox3.Text = sb.ToString();
        }

        private void appendText(string text) {
            this.Invoke((MethodInvoker)delegate {
                textBox3.Text += text;
            });
        }

        private void updateStatus(bool status) {
            this.Invoke((MethodInvoker)delegate {
                textBox1.Enabled = status;
                textBox2.Enabled = status;
                button1.Enabled = status;
                button2.Enabled = status;
                button3.Enabled = status;
                radioButton1.Enabled = status;
                radioButton2.Enabled = status;

            });
        }

    }
}
