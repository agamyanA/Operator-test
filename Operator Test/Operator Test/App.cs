using System;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace Operator_Test
{

    public partial class App : Form
    {
        double mxgum = 0, gum = 0;
        int[] answereSum = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        int amount = 0;
        double[] mx = { 0, 0, 0, 0, 0, 0, 0 };
        int[,] answere = new int[20, 20];
        double[] m = { 0, 0, 0, 0, 0, 0, 0 };
        int l = 0;
        static int i, j, rowCount, colCount, sheetIndex;
        static string[,] smas = new string[100, 10];
        int k, bx;
        string[] questions =
        {
             "Մասնագիտական հմտությունները",
             "Ոչ ստանդարտ իրավիճակներում որոշումներ կայացնելու ունակությունը",
             "Կազմակերպվածությունը",
             "Ուշադրությունը",
             "Հակվածությունը խմիչքի նկատմամբ",
             "Հոգեվիճակի կայությունությունը",
             "Ինտելեկտուալ զարգացվածության մակարդակը",
             "Սթրեսակայունությունը",
             "Ռեակցիայի արագությունը",
             "Հոգեկան հավասարակշռվածությունը",
             "Աշխատանքի ճշգրտությունը",
             "Աշխատանքային արտաքին միջավայրը",
             "Տեխնիկական և ծրագրային միջոցների հագեցվածությունը և արդիականությունը",
        };

        public static void getExcelFile(int sheet)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Application.StartupPath + @"\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
            sheetIndex = xlRange.Worksheet.Index;

            for (i = 1; i <= rowCount; i++)
            {
                for (j = 1; j <= 10; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        smas[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();
                }
            }

            xlApp.Workbooks.Close();

        }

        public void result_in_excel()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Application.StartupPath + @"\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[8];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            rowCount = xlWorksheet.Cells[xlWorksheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row + 1;

            xlWorksheet.Cells[rowCount, 1] = lastname_box.Text + " " + name_box.Text;
            xlWorksheet.Cells[rowCount, 2] = " " + m[bx];
            xlWorksheet.Cells[rowCount, 3] = " " + m[3];
            xlWorksheet.Cells[rowCount, 4] = " " + m[4];
            xlWorksheet.Cells[rowCount, 5] = " " + m[5];
            xlWorksheet.Cells[rowCount, 6] = " " + m[6];
            xlWorksheet.Cells[rowCount, 7] = " " + gum;

            xlWorkbook.SaveAs(Application.StartupPath + @"\data.xlsx", Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlApp.Workbooks.Close();
            xlApp.Quit();
        }

        public App()
        {
            InitializeComponent();
            welcomeLabel_2.Text = "Բարի գալուստ օպերատորների գնահատման թեստ";
            professionChoice_label.Text = "Ընտրեք Ձեր մասնագիտությունը";
        }

        void QandA(int i)
        {
            question_Label.Text = smas[i, 0];
            answereButton_1.Text = smas[i, 1];
            answereButton_2.Text = smas[i, 3];
            answereButton_3.Text = smas[i, 5];
        }

        void start()
        {
            picmain.Visible = false;
            name_box.Visible = false;
            lastname_box.Visible = false;
            signin_label.Visible = false;
            professionChoice_label.Visible = false;
            welcomeLabel_2.Visible = false;
            profession_boxChoice.Visible = false;
            startButton_2.Visible = resultLabel_3.Visible = false;
            nextButton_1.Visible = true;
            nextButton_1.Text = "Հաջորդ";
            nextButton_2.Visible = true;
            nextButton_2.Enabled = false;
            nextButton_2.Text = "Հաջորդ հարցերի խումբ";
            discribtionLabel_2.Visible = true;
            question_Label.Visible = true;
            answereButton_1.Visible = true;
            answereButton_2.Visible = true;
            answereButton_3.Visible = true;
            againButton_2.Visible = false;
            QandA(0);
            l = 0;
        }

        void count(int sht)
        {
            if (answereButton_1.Checked)
            {
                m[sht] = m[sht] + Convert.ToDouble(smas[l, 2]);
            }

            if (answereButton_2.Checked)
            {
                m[sht] = m[sht] + Convert.ToDouble(smas[l, 4]);
            }

            if (answereButton_3.Checked)
            {
                m[sht] = m[sht] + Convert.ToDouble(smas[l, 6]);
            }

            mx[sht] += Math.Max(Convert.ToDouble(smas[l, 2]), Math.Max(Convert.ToDouble(smas[l, 4]), Convert.ToDouble(smas[l, 6])));
            mxgum = mx[bx] + mx[3] + mx[4] + mx[5] + mx[6];
        }

        void pictures()
        {
            if (sheetIndex <= 3)
            {
                picprof.Visible = false;
                picatt.Visible = true;
                discribtionLabel_2.Text = "Որքան ուշադիր եք Դուք";
            }

            if (sheetIndex == 4)
            {
                picprof.Visible = false;
                picatt.Visible = false;
                picdec.Visible = true;
                discribtionLabel_2.Text = "Կարողանում եք Դուք արդյոք որոշումներ կայացնել";
            }

            if (sheetIndex == 5)
            {
                picdec.Visible = false;
                piciq.Visible = true;
                discribtionLabel_2.Text = "Ստուգենք Ձեր տրամաբանությունը";
            }

            if (sheetIndex == 6)
            {
                piciq.Visible = false;
                picpsyco.Visible = true;
                discribtionLabel_2.Text = "Եկեք իմանանք հավասարակշռված եք դուք թե ոչ";
            }
        }

        private void name_box_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void name_box_MouseDown(object sender, MouseEventArgs e)
        {
            name_box.Text = "";
            name_box.ForeColor = Color.Black;
        }

        private void lastname_box_MouseDown(object sender, MouseEventArgs e)
        {
            lastname_box.Text = "";
            lastname_box.ForeColor = Color.Black;
        }

        private void lastname_box_TextChanged(object sender, EventArgs e)
        {
            profession_boxChoice.Enabled = true;
        }

        private void profession_boxChoice_SelectedIndexChanged(object sender, EventArgs e)
        {
            startButton_2.Enabled = true;
            bx = profession_boxChoice.SelectedIndex;
            for (k = 0; k <= 2; k++)
                if (profession_boxChoice.SelectedIndex == k)
                    getExcelFile(k + 1);
        }

        private void startButton_2_Click(object sender, EventArgs e)
        {
            start();
            picprof.Visible = true;
            discribtionLabel_2.Text = "Ստուգենք Ձեր մասնագիտական գիտելիքները";
        }

        private void answereButton_1_MouseClick(object sender, MouseEventArgs mouseEvent)
        {
            if (l < rowCount - 1)
            {
                nextButton_1.Enabled = true;
            }
            else
                nextButton_2.Enabled = true;
        }

        private void nextButton_1_Click(object sender, EventArgs e)
        {
            nextButton_1.Enabled = false;
            count(sheetIndex - 1);
            answereButton_1.Checked = answereButton_2.Checked = answereButton_3.Checked = false;
            if (l < rowCount - 1)
            {
                l++;
                QandA(l);
            }
        }

        private void nextButton_2_Click(object sender, EventArgs e)
        {
            pictures();
            nextButton_2.Enabled = false;
            if (sheetIndex < 3)
                count(bx);
            else
                count(sheetIndex - 1);
            answereButton_1.Checked = answereButton_2.Checked = answereButton_3.Checked = false;
            if (sheetIndex < 7)
            {
                k++;
                getExcelFile(k);
                start();
                if (sheetIndex == 7)
                {
                    nextButton_2.Text = "Տեսնել արդյունքը";
                }
            }
            else
            {
                picpsyco.Visible = false;
                discribtionLabel_2.Visible = false;
                question_Label.Visible = false;
                answereButton_1.Visible = false;
                answereButton_2.Visible = false;
                answereButton_3.Visible = false;
                resultLabel_3.Visible = nextButton_1.Enabled = true;
                resultLabel_3.Text = "Արդյունք" + "\n\n\n";
                gum = m[bx] + m[3] + m[4] + m[5] + m[6];
                resultLabel_1.Visible = true;
                resultLabel_2.Visible = true;
                resultLabel_1.Text = " Մասնագիտական գիտելիքներ ";
                resultLabel_1.Text += "\n Ուշադրություն ";
                resultLabel_1.Text += "\n Որոշումներ կայացնելու ունակություն ";
                resultLabel_1.Text += "\n Ինտելեկտուալ զարգացվածություն ";
                resultLabel_1.Text += "\n Հոգեվիճակի կայունություն ";
                resultLabel_1.Text += "\n\n Ընդհանուր գնահատական ";
                resultLabel_2.Text = "   " + m[bx].ToString() + " / " + mx[bx].ToString() + "   " + (m[bx] / mx[bx] * 100).ToString("F1") + "%";
                resultLabel_2.Text += "\n   " + m[3].ToString() + " / " + mx[3].ToString() + "   " + (m[3] / mx[3] * 100).ToString("F1") + "%";
                resultLabel_2.Text += "\n   " + m[4].ToString() + " / " + mx[4].ToString() + "   " + (m[4] / mx[4] * 100).ToString("F1") + "%";
                resultLabel_2.Text += "\n   " + m[5].ToString() + " / " + mx[5].ToString() + "   " + (m[5] / mx[5] * 100).ToString("F1") + "%";
                resultLabel_2.Text += "\n   " + m[6].ToString() + " / " + mx[6].ToString() + "   " + (m[6] / mx[6] * 100).ToString("F1") + "%";
                if (gum > 9)
                    resultLabel_2.Text += "\n\n " + gum.ToString() + " / " + mxgum.ToString() + "   " + (gum / mxgum * 100).ToString("F1") + "%";
                else
                    resultLabel_2.Text += "\n\n    " + gum.ToString() + " / " + mxgum.ToString() + "   " + (gum / mxgum * 100).ToString("F1") + "%";
                result_in_excel();
                nextButton_1.Visible = false;
                nextButton_2.Visible = false;
                againButton_2.Visible = true;
            }
        }

        private void againButton_2_Click(object sender, EventArgs e)
        {
            mxgum = 0;
            Array.Clear(mx, 0, 7);
            Array.Clear(m, 0, 7);
            profession_boxChoice.SelectedIndex = -1;
            name_box.ResetText();
            lastname_box.ResetText();
            name_box.ForeColor = Color.FromKnownColor(KnownColor.InactiveCaption);
            lastname_box.ForeColor = Color.FromKnownColor(KnownColor.InactiveCaption);
            name_box.Text = "Անուն";
            lastname_box.Text = "Ազգանուն";
            picpsyco.Visible = false;
            profession_boxChoice.Enabled = false;
            picmain.Visible = true;
            name_box.Visible = true;
            lastname_box.Visible = true;
            signin_label.Visible = true;
            againButton_2.Visible = false;
            resultLabel_3.Visible = false;
            professionChoice_label.Visible = true;
            resultLabel_2.Visible = false;
            resultLabel_1.Visible = false;
            profession_boxChoice.Visible = true;
            startButton_2.Visible = true; startButton_2.Enabled = false;
        }

        void propertyLabel_text()
        {
            propertyLabel_1.Text = questions[0];
            propertyLabel_2.Text = questions[1];
            propertyLabel_3.Text = questions[2];
            propertyLabel_4.Text = questions[3];
            propertyLabel_5.Text = questions[4];
            propertyLabel_6.Text = questions[5];
            propertyLabel_7.Text = questions[6];
            propertyLabel_8.Text = questions[7];
            propertyLabel_9.Text = questions[8];
            propertyLabel_10.Text = questions[9];
            propertyLabel_11.Text = questions[10];
            propertyLabel_12.Text = questions[11];
            propertyLabel_13.Text = questions[12];
        }

        void rateCB_null()
        {
            rateCB_1.SelectedIndex = -1;
            rateCB_2.SelectedIndex = -1;
            rateCB_3.SelectedIndex = -1;
            rateCB_4.SelectedIndex = -1;
            rateCB_5.SelectedIndex = -1;
            rateCB_6.SelectedIndex = -1;
            rateCB_7.SelectedIndex = -1;
            rateCB_8.SelectedIndex = -1;
            rateCB_9.SelectedIndex = -1;
            rateCB_10.SelectedIndex = -1;
            rateCB_11.SelectedIndex = -1;
            rateCB_12.SelectedIndex = -1;
            rateCB_13.SelectedIndex = -1;
        }

        private void App_Load(object sender, EventArgs e)
        {

        }

        private void startButton_1_Click(object sender, EventArgs e)
        {
            propertyLabel_text();
            panel_rateCB.Visible = true;
            picmain_2.Visible = false;
            startButton_1.Visible = false;
            wellcomeLabel_1.Visible = false;
            panel_propertyLabel.Visible = true;
            discribtionLabel_1.Visible = true;
            count_Button.Visible = true;
        }

        private void rateCB_13_SelectedIndexChanged(object sender, EventArgs e)
        {
            count_Button.Enabled = true;
        }

        private void count_Button_Click(object sender, EventArgs e)
        {
            answere[0, amount] = rateCB_1.SelectedIndex + 1;
            answere[1, amount] = rateCB_2.SelectedIndex + 1;
            answere[2, amount] = rateCB_3.SelectedIndex + 1;
            answere[3, amount] = rateCB_4.SelectedIndex + 1;
            answere[4, amount] = rateCB_5.SelectedIndex + 1;
            answere[5, amount] = rateCB_6.SelectedIndex + 1;
            answere[6, amount] = rateCB_7.SelectedIndex + 1;
            answere[7, amount] = rateCB_8.SelectedIndex + 1;
            answere[8, amount] = rateCB_9.SelectedIndex + 1;
            answere[9, amount] = rateCB_10.SelectedIndex + 1;
            answere[10, amount] = rateCB_11.SelectedIndex + 1;
            answere[11, amount] = rateCB_12.SelectedIndex + 1;
            answere[12, amount] = rateCB_13.SelectedIndex + 1;
            amount++;
            count_Button.Enabled = true;
            viewresult.Visible = true;
            for (int i = 0; i < 13; i++)
                for (int j = 0; j < amount; j++)
                {
                    answereSum[i] += answere[i, j];
                }
            Array.Sort(answereSum, questions);
            panel_propertyLabel.Visible = false;
            panel_rateCB.Visible = false;
            discribtionLabel_1.Visible = false;
            count_Button.Visible = false;
            propertyLabel_text();
            againButton_1.Visible = true;
            viewresult.RowCount = 13;
            viewresult.ColumnCount = 3;
            for (int i = 0; i < 13; i++)
            {
                viewresult.Rows[i].Cells[0].Value = i + 1;
                viewresult.Rows[i].Cells[1].Value = questions[i];
                viewresult.Rows[i].Cells[2].Value = answereSum[i];
                viewresult.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                viewresult.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft;
                viewresult.Columns[1].DefaultCellStyle.Font = new Font("Sylfaen", 10, FontStyle.Regular);
                viewresult.Rows[i].Height = 24;
                viewresult.ClearSelection();
            }
            for (int i = 0; i < 5; i++)
            {
                viewresult.Rows[i].Cells[1].Style.ForeColor = Color.Red;
            }
        }

       private void againButton_1_Click(object sender, EventArgs e)
        {
            count_Button.Enabled = false;
            rateCB_null();
            picmain_2.Visible = true;
            viewresult.Visible = false;
            againButton_1.Visible = false;
            panel_propertyLabel.Visible = false;
            for (int i = 0; i < 13; i++)
                for (int j = 0; j < 13; j++)
                {
                    answereSum[i] = 0;
                    answere[i, j] = 0;
                }
            wellcomeLabel_1.Visible = true;
            startButton_1.Visible = true;
            rateCB_null();
            count_Button.Enabled = false;
        }

    }
}