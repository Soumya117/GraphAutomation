using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using Microsoft.Office.Core;
using System.Reflection;
using System.Web;

using Excel = Microsoft.Office.Interop.Excel;

namespace datagrapgh
{
   
    public partial class Form1 : Form
    {
       //private static double ri, ro, pi, po, r, ts, ls, rs, p1, p2, p3, p4,ls1,ls2,ts1,ts2,rs1,rs2i,j,k;
       private static double i,j,k;
      //private static string file;
       string path;
       Image image;
//Boolean b = true;
        DataTable dt = new DataTable();
        DataRow dr;
        public Form1()
        {
            InitializeComponent();
            // dt.Columns.Add("Inside Dia");
            // dt.Columns.Add("Outside Dia");
            //  dt.Columns.Add("Pressure on OD");
            //  dt.Columns.Add("Pressure on ID");
            dt.Columns.Add("Radius(r)");
            // dt.Columns.Add("Ends Capped");
            dt.Columns.Add("Tangential Stress");
            dt.Columns.Add("Longitudinal Stress");
            dt.Columns.Add("Radial Stress");
            // dt.Columns.Add("Created On");
            comboBox1.Items.Add("Yes");
            comboBox1.Items.Add("No");
            button4.Visible = false;
            label10.Text = DateTime.Now.ToString();
            monthCalendar1.Visible = false;
            button3.Visible = false;
            toolTip1.SetToolTip(this.label10, "Click to display the calendar");
            toolTip2.SetToolTip(this.button1, "Click to compute the results");
            toolTip1.SetToolTip(this.button2, "Click to reset the textboxes");
            SetStyle(ControlStyles.AllPaintingInWmPaint |
    ControlStyles.DoubleBuffer |
    ControlStyles.ResizeRedraw |
    ControlStyles.UserPaint,
    true);

        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.C))
            {
                button1_Click(null, null);
            }
            if (keyData == (Keys.Control | Keys.I))
            {
                button4_Click(null, null);
            }
            if (keyData == (Keys.Control | Keys.N))
            {
                button2_Click(null, null);
            }
            if (keyData == (Keys.Control | Keys.D))
            {
                label10_Click(null, null);
            }
            if (keyData == (Keys.Control | Keys.H))
            {
                button3_Click(null, null);
            }
            if (keyData == (Keys.Control | Keys.F1))
            {
                toolStripMenuItem1.ShowDropDown();
                //  toolStripMenuItem1.HideDropDown();
            }
            if (keyData == (Keys.Control | Keys.V))
            {
                toolStripMenuItem5.ShowDropDown();
                // toolStripMenuItem5.HideDropDown();
            }
            if (keyData == (Keys.Control | Keys.R))
            {
                toolStripMenuItem6.PerformClick();
            }
            if (keyData == (Keys.Control | Keys.S))
            {
                shortCutKeysToolStripMenuItem.PerformClick();
            }
            if (keyData == (Keys.Control | Keys.E))
            {
                toolStripMenuItem4.PerformClick();
                // toolStripMenuItem5.HideDropDown();
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }



        private void button1_Click(object sender, EventArgs e)
        {
            int interval;
            interval = int.Parse(numericUpDown1.Value.ToString());

            double ri, ro, pi, po, r, ts, ls, rs, p1, p2, p3, p4, ls1, ls2, ts1, ts2, rs1, rs2;
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox3.Text) || string.IsNullOrEmpty(textBox4.Text) || string.IsNullOrEmpty(textBox5.Text) || comboBox1.Text == "Select")
            {
                MessageBox.Show("Some fields are empty(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            { //double ri, ro, pi, po, r, ts, ls, rs, p1, p2, p3, p4;
               

                ri = double.Parse(textBox1.Text);
                //  double ri1 = Math.Round(ri, 2);
                ro = double.Parse(textBox2.Text);
                pi = double.Parse(textBox3.Text);
                po = double.Parse(textBox4.Text);
                r = double.Parse(textBox5.Text);

                if (ri < 0 || ro < 0 || pi < 0 || po < 0 || r < 0)
                {
                    MessageBox.Show("Values should be greater than zero(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (ri >= ro)
                    {
                        MessageBox.Show("Outside diameter shall be greater than or equal to inside diameter\n(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        p1 = (pi * ri * ri) - (po * ro * ro);
                        p2 = (ro * ro) - (ri * ri);
                        p3 = (ri * ri * ro * ro) * (po - pi);
                        p4 = (r * r) * (ro * ro - ri * ri);
                        ts = (p1 / p2) - (p3 / p4);
                         ts1 = Math.Round(ts, 3);
                        
                        textBox6.Text = ts1.ToString();
                        if (comboBox1.Text == "Yes")
                        {

                            ls2 = p1 / p2;
                             ls = Math.Round(ls2, 3);
                            // double ls2 = Math.Round(ls, 3);
                            textBox7.Text = ls.ToString();
                        }
                        else
                        {
                            ls = 0;
                            textBox7.Text = ls.ToString();
                        }
                        rs = (p1 / p2) + (p3 / p4);
                         rs1 = Math.Round(rs, 3);
                     
                        textBox8.Text = rs1.ToString();
                        for (i = ri; i <= ro; i += interval)
                        {

                            p1 = (pi * ri * ri) - (po * ro * ro);
                            p2 = (ro * ro) - (ri * ri);
                            p3 = (ri * ri * ro * ro) * (po - pi);
                            p4 = (i * i) * (ro * ro - ri * ri);
                            double tsn = (p1 / p2) - (p3 / p4);
                            ts2 = Math.Round(tsn, 3);
                            double rsn = (p1 / p2) + (p3 / p4);
                            rs2 = Math.Round(rsn, 3);
                            dr = this.dt.NewRow();
                            this.dt.Rows.Add(dr);
                            while (j <= k)
                            {

                                dr[0] = i;
                                dr[1] = ts2;
                                dr[2] = ls;
                                dr[3] = rs2;

                                dataGridView1.DataSource = dt;

                                j++;

                            }
                            k = k + 4;
                            // double ls2 = Math.Round(ls, 3);
                            // double p41 = (i * i) * (ro * ro - ri * ri);
                            // double ts2 = (p1 / p2) - (p3 / p41);
                            //  ls2 = p1 / p2;
                            // double rs2 = Math.Round(rs, 3);
                            // rs2 = (p1 / p2) + (p3 / p41);

                        }
                       
                    }//else over

                }//else

                //for
            }//else 1st
        }//button
        private void button4_Click(object sender, EventArgs e)
        {

         //   dr = this.dt.NewRow();
            // dr[0] = textBox1.Text;
            // dr[1] = textBox2.Text;
            //dr[2] = textBox3.Text;
            //dr[3] = textBox4.Text;
            // dr[4] = textBox5.Text;
            //dr[5] = comboBox1.Text;
            // dr[6] = textBox6.Text;
            // dr[7] = textBox7.Text;
            // dr[8] = textBox8.Text;
         //   dr[0] = textBox5.Text;
         //   dr[1] = textBox6.Text;
         //   dr[2] = textBox7.Text;
          //  dr[3] = textBox8.Text;
            // dr[4] = DateTime.UtcNow.ToShortDateString().ToString();
         //   this.dt.Rows.Add(dr);
          //  dataGridView1.DataSource = dt;

            

        }
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "Select";
            textBox1.ResetText();
            textBox2.ResetText();
            textBox3.ResetText();
            textBox4.ResetText();
            textBox5.ResetText();
            textBox6.ResetText();
            textBox7.ResetText();
            textBox8.ResetText();

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("My software", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://www.google.com");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("An error occurred {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void onlineHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            String s;
            s = textBox1.Text;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Only Numbers(See the rules in the View Menu)", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox1.ResetText();
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            String s;
            s = textBox2.Text;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Only Numbers", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox2.ResetText();
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

            String s;
            s = textBox3.Text;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Only Numbers", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox3.ResetText();
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            String s;
            s = textBox4.Text;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Only Numbers", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox4.ResetText();
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

            String s;
            s = textBox5.Text;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Only Numbers", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox5.ResetText();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            monthCalendar1.Visible = false;
            button3.Visible = false;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle = cp.ExStyle | 0x2000000;
                return cp;
            }
        }

        /* public override Image BackgroundImage
         {
             get
             {
                 return base.BackgroundImage;
             }
             set
             {
                 if (value != null)
                 {
                     //Create a new bitmap image has same same size of the form
                     Bitmap m_bmp = new Bitmap(
                         this.Width,
                         this.Height,
                         System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                     //Make a graphics instance that can draw on our bitmap
                     Graphics g = Graphics.FromImage(m_bmp);

                     //Here is the magic
                     //We want to render our compressed image to a raw bitmap image
                     //This line is like uncompressing our image and scaling it
                     g.DrawImage(value, 0, 0, m_bmp.Width, m_bmp.Height);

                     //Don`t forget to release your resources
                     g.Dispose();

                     //Assign our bitmap to the base form
                     base.BackgroundImage = m_bmp;
                 }
                 else
                 {
                     base.BackgroundImage = null;
                 }
             }
         }*/
        private void label10_Click(object sender, EventArgs e)
        {
            monthCalendar1.Visible = true;
            button3.Visible = true;
        }

        private void rulesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("All fields should be numerical.\nAll boxes should be field.\nAll numbers should be greater than Zero.\nOutside diameter should be greater than inside diameter.\n", "Rules", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void contactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Contact At: \n Email ID: soumya113157@nitp.ac.in", "Contact", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void button5_Click(object sender, EventArgs e)
        {

            object misValue = System.Reflection.Missing.Value;

            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            // see the excel sheet behind the program
            app.Visible = true;
            try
            {
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                // changing the name of active sheet
                worksheet.Name = "Exported from gridview";


                // storing header part in Excel
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }



                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                // save the application
                string fileName = String.Empty;
                SaveFileDialog saveFileExcel = new SaveFileDialog();

                saveFileExcel.Filter = "Excel files |*.xls|All files (*.*)|*.*";
                saveFileExcel.FilterIndex = 2;
                saveFileExcel.RestoreDirectory = true;

                if (saveFileExcel.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileExcel.FileName;
                    //Fixed-old code :11 para->add 1:Type.Missing
                    workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                }
                else
                    return;

                workbook = app.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                //MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
                chartRange = worksheet.get_Range("A1", "D30");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.Export("D:\\Image.jpeg", "JPEG", misValue);
                Image image = Image.FromFile("D:\\Image.jpeg");
                pictureBox1.Image = image;
                // workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                //app.Quit();
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
                app.Quit();
                workbook = null;
                app = null;
            }
            // save the application

            //   workbook.SaveAs("D:\\csharp.net-informations2.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            // Exit from the application
            //app.Quit();
        }









     
        
        
        
        private void graph()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("D:\\csharp.net-informations2.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
            chartRange = xlWorkSheet.get_Range("A1", "D30");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.Export("D:\\Image.jpeg", "JPEG", misValue);
            Image image = Image.FromFile("D:\\Image.jpeg");
            pictureBox1.Image = image;
            //    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            dt.Clear();
        }



        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Software Version:1.0\nGraph Automation", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://www.google.com");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("An error occurred {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("All fields should be numerical.\nAll boxes should be field.\nAll numbers should be greater than Zero.\nOutside diameter should be greater than inside diameter.\n", "Rules", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Contact At: \n Email ID:", "Contact", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void shortCutKeysToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Compute(CTRL+C)\nInsert(CTRL+I)\nReset(CTRL+N)\nShow Calendar(CTRL+D)\nHide Calendar(CTRL+H)\nEXIT(CTRL+E)\nRules(CTRL+R)\nShortCut Keys(CTRL+S)\nHelp Menu(CTRL+F1)\nView Menu(CTRL+V)", "Access Keys", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

       

       

        private void button7_Click_1(object sender, EventArgs e)
        {
            string folderdate = DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss");
            
          

                path = "C:\\Export\\" + folderdate;
          

            if (!Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
                
            }

            string filename = path + "\\exprt.xls";
            string imagenew = path + "\\Image.jpeg";
            string exclimage = path + "\\exclimg.xls";
           
            object misValue = System.Reflection.Missing.Value;

            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            // see the excel sheet behind the program
            app.Visible = true;
          //  try
           // {
                // get the reference of first sheet. By default its name is Sheet1.
                // store its reference to worksheet
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                // changing the name of active sheet
                worksheet.Name = "Exported from gridview";


                // storing header part in Excel
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }



                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

             
              
                workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook = app.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                //MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
                chartRange = worksheet.get_Range("A1", "D1000");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.Export(imagenew, "JPEG", misValue);
                image = Image.FromFile(imagenew);
                pictureBox1.Image = image;
                workbook.SaveAs(exclimage, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
              // MessageBox.Show("All the Excel files and Images are stored successfully in C:\\Export", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                //app.Quit();
                workbook.Close(true, misValue, misValue);
                //app.Quit();

                //app.Quit();
               
          //  }
          //  catch (System.Exception ex)
         //  {
         //      MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
         //  }
        //  finally
          //  {
                app.Quit();
                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(app);
              
               workbook = null;
               app = null;
              // workbook.Close(true, misValue, misValue);
           //}

        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            //LoadNewFile();
            graphext();

        }
        private void graphext()
        {
            string folderdate = DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss");


         

                path = "C:\\Export\\" + folderdate;
           

            //string path1 = "C:\\Export\\" + folderdate + "\\"+"image.jpeg";
            //string path2 = "C:\\Export\\" + folderdate + "\\"+"excelnew.xls";

            if (!Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);

            }

            //string filename = path + "\\exprt.xls";
            string imagenew = path + "\\Image.jpeg";
            string exclimage = path + "\\exclimg.xls";

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
           // openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = String.Empty;
            openFileDialog1.Filter = "Excel Sheet(.xls)|*.xls|Microsoft Excel Sheets(.xlsx)|*.xlsx";
           // System.Windows.Forms.DialogResult dr = openFileDialog1.ShowDialog();
            //   string filename = openFileDialog1.FileName;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                xlApp = new Excel.Application();
                // xlWorkBook = xlApp.Workbooks.Open("D:\\Book1.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            }
            else
                return;
                //MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
                chartRange = xlWorkSheet.get_Range("A1", "D1000");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.Export(imagenew, "JPEG", misValue);
                 image = Image.FromFile(imagenew);
                pictureBox1.Image = image;
                //    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
               
           xlWorkBook.SaveAs(exclimage, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
           MessageBox.Show("All the Excel files and Images are stored successfully in C:\\Export", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            
        }
     
        private void button8_Click(object sender, EventArgs e)
        {

            //saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png|Tiff Image (.tiff)|*.tiff|Wmf Image (.wmf)|*.wmf";
           // System.Windows.Forms.DialogResult dr = saveFileDialog1.ShowDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Jpeg);
            }
            else
                return;
        }

        private void goToFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = "C:\\Export";
            DirectoryInfo dir = new DirectoryInfo(path);
            if (Directory.Exists(path))
            {
                System.Diagnostics.Process.Start("explorer.exe", @"C:\Export");
            }
            else
            {
                MessageBox.Show("Directory doesnt exist", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void manageSpaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Image.Dispose();

           pictureBox1.ResetText();
         //   releaseObject(image);
            
           pictureBox1.Image = null;
            if (MessageBox.Show("Delete all the files?", "Manage Space", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {


               // string folderdate = DateTime.Now.ToFileTime().ToString();

                //pictureBox1.Refresh();
                string path = @"C:\Export\";
                DirectoryInfo dir = new DirectoryInfo(path);

                if (Directory.Exists(path))
                {
                    
                    DeleteDirectory(path);
                    MessageBox.Show("Directory successfully deleted", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    MessageBox.Show("Folder Doesnt Exist", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
               
            }
            else
                return;
}
        

        private void DeleteDirectory(string path)
        {
           
           
            // Delete all files from the Directory
            foreach (string filename in Directory.GetFiles(path))
            {
                File.Delete(filename);
            }
            // Check all child Directories and delete files
            foreach (string subfolder in Directory.GetDirectories(path))
            {
                DeleteDirectory(subfolder);
            }
            Directory.Delete(path);
           
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png|Tiff Image (.tiff)|*.tiff|Wmf Image (.wmf)|*.wmf";
            // System.Windows.Forms.DialogResult dr = saveFileDialog1.ShowDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Jpeg);
            }
            else
                return;
        }

        private void openFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = "C:\\Export";
            DirectoryInfo dir = new DirectoryInfo(path);
            if (Directory.Exists(path))
            {
                System.Diagnostics.Process.Start("explorer.exe", @"C:\Export");
            }
            else
            {
                MessageBox.Show("Directory doesnt exist", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void exitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    pictureBox1.Image.Dispose();
        //}

      
    }
}

       
  
       
    

