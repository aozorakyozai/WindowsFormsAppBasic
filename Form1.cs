using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsAppBasic
{
    public partial class setting : System.Windows.Forms.Form
    {
        Excel.Application elxApp;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range rng;
        // コンストラクタ
        public setting()
        {
            InitializeComponent();
            MessageBox.Show("指定したフォルダからファイル名を取得する");
            // 指定してフォルダの中のファイルを探す
            string[] files = Directory.GetFiles(@"C:\testFolder\", "*");
            try
            {
                // 検索したファイルを表示する
                foreach (string f in files)
                {
                    MessageBox.Show(f);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void L1_Click(object sender, EventArgs e)
        {
            
            // Excel起動
            elxApp = new Excel.Application();
            // Excelのウィンドウを表示・非表示
            elxApp.Visible = true;
            elxApp.Application.DisplayAlerts = true;

            // Bookを追加
            //elxapp.Application.Workbooks.Add(Type.Missing);
            //csWorkbook = new Excel.Workbook();

            /******** Excel WorkBook ********/
            // C直下にtestFolderを作成し、testSc.xlsを作成する
            // パスを指定してエクセルを開く
            // Type.Missing：特に指定しない場合
            wb = elxApp.Workbooks.Open(@"C:\testFolder\testCs", Type.Missing);

            /******** Sheet ********/

            // Bookを追加
            //Excel.Workbook wb;
            //wb = csWorkbook.Application.Workbooks.Add(Type.Missing);

            // シート選択
            ws = wb.Worksheets[2];

            // セルの値を取得
            rng = ws.Cells[1, 1];
            MessageBox.Show(rng.Value);

            // Formを終了する
            //Application.Exit();
        }

        private void R1_Click(object sender, EventArgs e)
        {
            // Excel保存
            //wb.SaveAs(@"C:\testFolder\testCs");
        }
        private void Form1_Clik(object sender, EventArgs e)
        {
            
        }
        

        
    }
}
