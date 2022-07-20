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
using ClosedXML.Excel;

namespace setting
{
    public partial class setting : System.Windows.Forms.Form
    {
        public DateTime today = DateTime.Today;
        public DateTime now = DateTime.Now;

        /**************** コンストラクタ ****************/
        public setting()
        {
            InitializeComponent();
        }

        /**************** Button ****************/

        // シート作成 ＞ F5に値を入力する > 保存
        public void L1_Click(object sender, EventArgs e)
        {
            try
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Tab test02");
                worksheet.Cell(5, 5).Value = "test";
                worksheet.Cell(5, 6).Value = "Hello world!";
                worksheet.Cell("F5").Style.Font.FontSize = 24; // 5,6と同じ
                workbook.SaveAs(@"C:\Users\iwanami\MyDrive\myNote.xlsx");
            }
             catch
            {
                MessageBox.Show("Err L1");
            }
        }
        // F5の値を取得する
        private void L2_Click(object sender, EventArgs e)
        {
            try
            {
                var workbook = new XLWorkbook(@"C:\Users\iwanami\MyDrive\myNote.xlsx");
                foreach (var worksheet in workbook.Worksheets)
                {
                    // シートを探す
                    if (worksheet.Name.Equals("Tab test02"))
                    {
                        // セルの値を取得する
                        Console.WriteLine("F5 = {0}", worksheet.Cell("F5").Value);
                        break;
                    }
                }
            }
            catch
            {
                MessageBox.Show("Err L2");
            }
        }

        // 検索
        // 検索内容：シート名&＆　カラム数&＆　一つ前のセルの値が「test」の場合のセルの値を取得する（複数セルが可能）
        private void L3_Click(object sender, EventArgs e)
        {
            try
            {
                var workbook = new XLWorkbook(@"C:\Users\iwanami\MyDrive\myNote.xlsx");
                var cells = workbook.FindCells((cell) => 
                {
                    var addr = cell.Address;
                    // StartsWith(---): ---から始まる文字列 cf:EndsWith
                    if (cell.Worksheet.Name.Equals("Tab test02")
                    && addr.ColumnNumber == 6
                    && cell.Worksheet.Cell(addr.RowNumber, addr.ColumnNumber -1).Value.ToString().StartsWith("te"))
                    {
                        return true;
                    }
                    return false;
                });
                foreach (var cell in cells)
                {
                    Console.WriteLine("{0} = {1}", cell.Address, cell.Value);
                }
            }
            catch
            {
                Console.WriteLine("Err L3");
            } 
        }

        // 文字列の加工
        private void L4_Click(object sender, EventArgs e)
        {

            try
            {
                var workbook = new XLWorkbook(@"C:\Users\iwanami\MyDrive\myNote.xlsx");
                //var worksheet = workbook.Worksheets.Add("Tab test01");
                var worksheet2 = workbook.Worksheet("Tab test01");
                // シート名を指定　上記と同じ
                var worksheet1 = workbook.Worksheet(1);
                worksheet1.Cell(1, 7).Value = now;
                worksheet2.Range(1, 1, 10, 1).Value = "あ";
                // 上記と同じ入力方法
                worksheet2.Range(1, 2, 10, 2).SetValue("い");
                worksheet2.Range(1, 1, 10, 2).Style.Fill.BackgroundColor = XLColor.YaleBlue;
                worksheet2.Cell(6, 6).Style.NumberFormat.Format = "yyyy/m/d";
                worksheet2.Cell(6, 6).Value = today + " This is Today test 03"; // 今日の日付を入力
                workbook.SaveAs(@"C:\Users\iwanami\MyDrive\myNote.xlsx");
            }
            catch
            {
                Console.WriteLine("Err L4");
            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

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
            dr["id"] = "09";
            dr["name"] = "Close";
            dt.Rows.Add(dr);

            this.comboBox1.DataSource = dt;
            this.comboBox1.DisplayMember = "Name";
            this.comboBox1.ValueMember = "id";
            this.comboBox1.ItemHeight = 45; // コンボボックスの高さ
            this.comboBox1.DropDownWidth = 200; // メニューの幅

            /******** ToolTipの設定 ********/
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000; // ５秒 表示時間(ミリ秒)
            toolTip1.InitialDelay = 100; // 待機時間
            toolTip1.ReshowDelay = 100; // 次のヒントの待機時間
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(this.L1, "Make");
            toolTip1.SetToolTip(this.L2, "取得");
            toolTip1.SetToolTip(this.L3, "検索");
            toolTip1.SetToolTip(this.L4, "未定");
            toolTip1.SetToolTip(this.checkBox1, "Excel 終了");
            toolTip1.SetToolTip(this.checkBox4, "終了");
        } 
        private void setting_Closing(object sender, EventArgs e)
        {
            MessageBox.Show("xxxx");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

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
            if (radioButtonCenter.Checked == true)
            {
                MessageBox.Show("radiobutton Centerをクリックしました");
            }
            if (radioButtonUp.Checked == true)
            {
                MessageBox.Show("radiobutton Upをクリックしました");
            }
            if (radioButtonDown.Checked == true)
            {
                MessageBox.Show("radiobutton Downをクリックしました");
            }
            if (radioButtonLeft.Checked == true)
            {
                MessageBox.Show("radiobutton Leftをクリックしました");
            }
            if (radioButtonRight.Checked == true)
            {
                MessageBox.Show("radiobutton をクリックしました");
            }
        }

        /**************** checkBox ****************/

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            
            if (checkBox4.Checked == true)
            {
                MessageBox.Show("Close");
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
        private object workbook;

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

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

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
