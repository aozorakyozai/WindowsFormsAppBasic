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
    public partial class setting : Form
    {
        Excel.Application elxApp;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range rng;
        // コンストラクタ
        public setting()
        {
            InitializeComponent(); // フォーム作成
        }
        private void setting_Closing(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        /**************** Button ****************/
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
            // シート名の取得
            MessageBox.Show(elxApp.ActiveSheet.Name + " 【シート名を取得】");
            // セルの値を取得
            try
            {
                rng = ws.Cells[1, 1];

                MessageBox.Show(rng.Value + "　【セルの値】");
            }
            catch(ArgumentNullException ex)
            {
                MessageBox.Show(ex.Message + "　【セルの値の取得失敗】");
            }
            
        }
        private void L2_Click(object sender, EventArgs e)
        {
            // Excel保存
            wb.Save();
            // 別名で保存
            //this.book.SaveCopyAs(newFileName);
            this.wb.Close(false);
            // エクセルを閉じる
            this.elxApp.Quit();
        }

        private void L3_Click(object sender, EventArgs e)
        {
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
            // MySql 接続
            MySqlControl mySqlControl = new MySqlControl();
            if (mySqlControl.connect() == true)
            {
                MessageBox.Show("接続成功！");
                mySqlControl.deconnect();
            }
            else
            {
                MessageBox.Show("接続に失敗しました");
            }
        }

        private void L4_Click(object sender, EventArgs e)
        {
            // 検索
            var keyword = textBox1.Text;
            try
            {
                var hitRange = ws.Cells.Find(What: keyword, LookIn: -4163, LookAt: 1);
                if (hitRange != null)
                {
                    hitRange.Select();
                }
                else
                {
                    MessageBox.Show("No Keyword");
                }
                // 置換
                var sampleReplace = keyword.Replace(keyword, keyword + "01");
                textBox1.Text = sampleReplace;
                hitRange.Value = sampleReplace;
            }
            catch
            {
                MessageBox.Show("Err");
            }
        }

        private void setting_Load(object sender, EventArgs e)
        {
            /******** combBox 表示リストの設定 ********/
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");

            DataRow dr = dt.NewRow();
            dr["id"] = "";
            dr["name"] = "";
            dt.Rows.Add(dr); 

            dr = dt.NewRow();
            dr["id"] = "01";
            dr["name"] = "Leaning";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["id"] = "02";
            dr["name"] = "Edit";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["id"] = "03";
            dr["name"] = "Test";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["id"] = "04";
            dr["name"] = "Setting";
            dt.Rows.Add(dr);

            this.comboBox1.DataSource = dt;
            this.comboBox1.DisplayMember = "Name";
            this.comboBox1.ValueMember = "id";
            this.comboBox1.ItemHeight = 31; // コンボボックスの高さの変更

            /******** ToolTipの設定 ********/
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000; // ５秒 表示時間(ミリ秒)
            toolTip1.InitialDelay = 100; // 待機時間
            toolTip1.ReshowDelay = 100; // 次のヒントの待機時間
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(this.L2, "Save");
            toolTip1.SetToolTip(this.checkBox1, "Excel 終了");
            toolTip1.SetToolTip(this.checkBox4, "終了");
        }
        
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            int index = this.comboBox1.SelectedIndex;
            string value = this.comboBox1.SelectedValue.ToString();
            MessageBox.Show(value);
            return;
        }

        /**************** radioButton ****************/
        private void RadioButton_Checked(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                MessageBox.Show("radiobutton１をクリックしました");
            }
        }

        /**************** checkBox ****************/
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                this.Close();
            } 
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {

            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                /******* Random *******/
                var rand = new Random();
                int r = rand.Next(1, 9);
                MessageBox.Show(r.ToString());
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                // Excel保存
                wb.Save();
                // 別名で保存
                //this.book.SaveCopyAs(newFileName);
                this.wb.Close(false);
                // エクセルを閉じる
                this.elxApp.Quit();
            }
        }
        
        /**************** tab ****************/
        /*
         * TabPages 〉[…]からtab名などを変更できる
         */
        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        /**************** mouse ****************/
        //マウスのクリック位置を記憶
        private Point mousePoint;

        //Form1のMouseDownイベントハンドラ
        //マウスのボタンが押されたとき
        private void Form1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                //位置を記憶する
                mousePoint = new Point(e.X, e.Y);
            }
        }

        //Form1のMouseMoveイベントハンドラ
        //マウスが動いたとき
        private void Form1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                this.Left += e.X - mousePoint.X;
                this.Top += e.Y - mousePoint.Y;
                //または、つぎのようにする
                //this.Location = new Point(
                //    this.Location.X + e.X - mousePoint.X,
                //    this.Location.Y + e.Y - mousePoint.Y);
            }
        }
    }
}
/********* Windows Form Namual ********/
/*
 * イベント発生：雷マーク 〉発生するイベントを選択 〉ビルド
 * 文字列の改行1：\r\n
 * 文字列の改行2：文字列を直接改行してコードを書く(文字列は変数にする)
 * size自動調整：SpritContainer(縦横に分割できる)
 */ 