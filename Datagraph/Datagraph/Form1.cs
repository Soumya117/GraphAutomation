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
using System.Reflection;
using System.Web;

namespace datagraph
{
  public partial class PlotGraph : Form
  {
    string selectedPath;
    Boolean isDefaultLoc = true;
    Boolean isRoot = true;
    DataTable dataTable = null;
    ExcelObject excelObject = new ExcelObject();

    public PlotGraph()
    {
      InitializeComponent();
      populateInitialValues();
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
      if (keyData == (Keys.Control | Keys.C))
      {
        button1_Click(null, null);
      }
      if (keyData == (Keys.Control | Keys.N))
      {
        button2_Click(null, null);
      }
      if (keyData == (Keys.Control | Keys.D))
      {
        label10_Click(null, null);
      }
      if (keyData == (Keys.Control | Keys.F1))
      {
        toolStripMenuItem1.ShowDropDown();
      }
      if (keyData == (Keys.Control | Keys.V))
      {
        toolStripMenuItem5.ShowDropDown();
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
      }
      return base.ProcessCmdKey(ref msg, keyData);
    }

    private void button1_Click(object sender, EventArgs e)
    {
      calculateResults();
    }

    private void button2_Click(object sender, EventArgs e)
    {
      if (checkBox1.Checked)
      {
        reset();
      }
      else
      {
        comboBox1.Text = "Select";
        FormUtil.clearTextBoxes(Controls);
        numericUpDown1.Value = 0;
      }
    }

    public void InvokeMethod(Delegate method, params object[] args)
    {
      resetGraphImage();
      tabControl1.SelectedIndex = 2;
      label21.Visible = true;
      label21.Text = "Plotting Graph..Please wait..";
      excelObject.setDirectoryFlags(isDefaultLoc, isRoot, selectedPath);

      method.DynamicInvoke(args);

      label21.Visible = false;
      pictureBox1.Image = excelObject.getImage();
    }

    private void plotGraph(params object[] args)
    {
      resetGraphImage();
      tabControl1.SelectedIndex = 2;
      label21.Visible = true;
      label21.Text = "Plotting Graph..Please wait..";
      excelObject.setDirectoryFlags(isDefaultLoc, isRoot, selectedPath);
      var dataTable = (DataTable)args[0];
      if (args[1] is string)
      {
        var varparam = (string)args[1];
        excelObject.plotGraph(dataTable, varparam);
      }
      else
      {
        excelObject.plotGraph(dataTable, (DataGridView)args[1]);
      }
      label21.Visible = false;
      pictureBox1.Image = excelObject.getImage();
    }

    private void button5_Click_1(object sender, EventArgs e)
    {
      try
      {
        openFileDialog1.FileName = String.Empty;
        openFileDialog1.Filter = "Excel Sheet(.xls)|*.xls|Microsoft Excel Sheets(.xlsx)|*.xlsx";

        if (openFileDialog1.ShowDialog() == DialogResult.OK)
        {
          Invoke(new Action(() => plotGraph(dataTable, openFileDialog1.FileName)));
        }
      }
      catch (System.Exception ex)
      {
        MessageBox.Show(ex.ToString(), "Error during exporting data", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    private void button7_Click_1(object sender, EventArgs e)
    {
      var dataTableCount = System.Convert.ToInt32(dataTable.Rows.Count);
      if (dataTableCount == 0)
      {
        MessageBox.Show("No data points loaded");
        return;
      }
      Invoke(new Action(() => plotGraph(dataTable, dataGridView1)));
    }

    private void button8_Click(object sender, EventArgs e)
    {
      saveFileDialog1.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png|Tiff Image (.tiff)|*.tiff|Wmf Image (.wmf)|*.wmf";
      if (saveFileDialog1.ShowDialog() == DialogResult.OK)
      {
        pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Jpeg);
      }
      else
        return;
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
      FormUtil.verifyTextBoxInputIsNum(textBox1);
    }

    private void textBox2_TextChanged(object sender, EventArgs e)
    {
      FormUtil.verifyTextBoxInputIsNum(textBox2);
    }

    private void textBox3_TextChanged(object sender, EventArgs e)
    {
      FormUtil.verifyTextBoxInputIsNum(textBox3);
    }

    private void textBox4_TextChanged(object sender, EventArgs e)
    {
      FormUtil.verifyTextBoxInputIsNum(textBox4);
    }

    private void textBox5_TextChanged(object sender, EventArgs e)
    {
      FormUtil.verifyTextBoxInputIsNum(textBox5);
    }

    private void label10_Click(object sender, EventArgs e)
    {
      monthCalendar1.Visible = monthCalendar1.Visible ? false : true;
    }

    private void rulesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      MessageBox.Show("All fields should be numerical.\nAll boxes should be field.\nAll numbers should be greater than Zero.\nOutside diameter should be greater than inside diameter.\n", "Rules", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void contactToolStripMenuItem_Click(object sender, EventArgs e)
    {
      MessageBox.Show("Contact At: \n Email ID: soumya113157@nitp.ac.in", "Contact", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
    }

    private void toolStripMenuItem2_Click(object sender, EventArgs e)
    {
      string copyright = "\u00a9 Copyright 2014.";
      MessageBox.Show("Software Version:1.0.0.0\nGraph Automation\n\n\n" + copyright, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

    private void goToFolderToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string path = FormUtil.setFilePath(selectedPath, isDefaultLoc, isRoot);
      DirectoryInfo dir = new DirectoryInfo(path);
      if (Directory.Exists(path))
      {
        System.Diagnostics.Process.Start("explorer.exe", path);
      }
      else
      {
        MessageBox.Show("Directory doesnt exist", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
      }
    }

    private void manageSpaceToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string path = FormUtil.setFilePath(selectedPath, isDefaultLoc, isRoot);

      if (MessageBox.Show("Delete all the files?", "Manage Space", MessageBoxButtons.OKCancel) == DialogResult.OK)
      {
        if (pictureBox1.Image != null)
        {
          pictureBox1.Image.Dispose();
          pictureBox1.Image = null;
          pictureBox1.ResetText();
        }

        DirectoryInfo dir = new DirectoryInfo(path);

        if (Directory.Exists(path))
        {

          FormUtil.deleteDirectory(path);
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

    private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      if (pictureBox1.Image != null)
      {
        saveFileDialog1.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png|Tiff Image (.tiff)|*.tiff|Wmf Image (.wmf)|*.wmf";
        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
        {
          pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Jpeg);
        }
        else
          return;
      }
    }

    private void openFolderToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string path = FormUtil.setFilePath(selectedPath, isDefaultLoc, isRoot);
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

    private void changeFolderDestinationToolStripMenuItem_Click(object sender, EventArgs e)
    {
      if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
      {
        selectedPath = folderBrowserDialog2.SelectedPath;
        DirectoryInfo d = new DirectoryInfo(selectedPath);
        if (d.Parent == null)
        {
          isDefaultLoc = false;
          MessageBox.Show("This is a root folder", "Info", MessageBoxButtons.OK);
        }
        else
        {
          MessageBox.Show("Path is:" + selectedPath, "Info", MessageBoxButtons.OK);
          isRoot = false;
        }
      }
      else return;
    }

    private void defaultFolderToolStripMenuItem_Click(object sender, EventArgs e)
    {
      MessageBox.Show("Default folder is set to C:", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
      isDefaultLoc = true;
      isRoot = true;
    }

    private void resetToolStripMenuItem_Click(object sender, EventArgs e)
    {
      if (pictureBox1.Image != null)
      {
        pictureBox1.Image.Dispose();
        pictureBox1.Image = null;
        pictureBox1.ResetText();
      }
    }

    //////////////////////////////////    private user defined methods    //////////////////////////////////////////////

    protected override CreateParams CreateParams
    {
      get
      {
        CreateParams cp = base.CreateParams;
        cp.ExStyle = cp.ExStyle | 0x2000000;
        return cp;
      }
    }

    private void calculateResults()
    {
      DataRow dataRow;
      double
        inRadius,
        outRadius,
        idPressure,
        odPressure,
        stressRadius,
        longitudinalStress,
        tangentialStessText,
        stressRadText,
        radiusGrid = 0,
        tangentialStressGrid,
        radialStressGrid;

      dataTable.Clear();
      resetGraphImage();

      double steps = Double.Parse(numericUpDown1.Value.ToString());

      if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox3.Text) || string.IsNullOrEmpty(textBox4.Text) || string.IsNullOrEmpty(textBox5.Text) || comboBox1.Text == "Select" || numericUpDown1.Value == 0)
      {
        MessageBox.Show("Some fields are empty(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
      }
      else
      {
        inRadius = double.Parse(textBox1.Text);
        outRadius = double.Parse(textBox2.Text);
        idPressure = double.Parse(textBox3.Text);
        odPressure = double.Parse(textBox4.Text);
        stressRadius = double.Parse(textBox5.Text);

        if (inRadius < 0 || outRadius < 0 || idPressure < 0 || odPressure < 0 || stressRadius < 0)
        {
          MessageBox.Show("Values should be greater than zero(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
          if (inRadius >= outRadius)
          {
            MessageBox.Show("Outside diameter shall be greater than or equal to inside diameter\n(See the rules in the View Menu)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
          }
          else
          {
            double j = 0, k = 0;
            double interval;
            interval = outRadius / steps;
            double intrvl = Math.Round(interval, 4);
            var p1 = (idPressure * inRadius * inRadius) - (odPressure * outRadius * outRadius);
            var p2 = (outRadius * outRadius) - (inRadius * inRadius);
            var p3 = (inRadius * inRadius * outRadius * outRadius) * (odPressure - idPressure);
            var p4 = (stressRadius * stressRadius) * (outRadius * outRadius - inRadius * inRadius);
            var tangentialStress = (p1 / p2) - (p3 / p4);
            tangentialStessText = Math.Round(tangentialStress, 4);
            textBox6.Text = tangentialStessText.ToString();

            if (comboBox1.Text == "Yes")
            {
              var longitudinalStressTmp = p1 / p2;
              longitudinalStress = Math.Round(longitudinalStressTmp, 4);
              textBox7.Text = longitudinalStress.ToString();
            }
            else
            {
              longitudinalStress = 0;
              textBox7.Text = longitudinalStress.ToString();
            }
            var radialStress = (p1 / p2) + (p3 / p4);
            stressRadText = Math.Round(radialStress, 4);
            textBox8.Text = stressRadText.ToString();

            for (radiusGrid = inRadius; radiusGrid <= outRadius; radiusGrid += intrvl)
            {
              p1 = (idPressure * inRadius * inRadius) - (odPressure * outRadius * outRadius);
              p2 = (outRadius * outRadius) - (inRadius * inRadius);
              p3 = (inRadius * inRadius * outRadius * outRadius) * (odPressure - idPressure);
              p4 = (radiusGrid * radiusGrid) * (outRadius * outRadius - inRadius * inRadius);
              double tsn = (p1 / p2) - (p3 / p4);
              tangentialStressGrid = Math.Round(tsn, 4);
              double rsn = (p1 / p2) + (p3 / p4);
              radialStressGrid = Math.Round(rsn, 4);
              dataRow = this.dataTable.NewRow();
              this.dataTable.Rows.Add(dataRow);
              while (j <= k)
              {
                dataRow[0] = radiusGrid;
                dataRow[1] = tangentialStressGrid;
                dataRow[2] = longitudinalStress;
                dataRow[3] = radialStressGrid;
                dataGridView1.DataSource = dataTable;
                j++;
              }
              k = k + 4;
            }
          }
        }
      }
    }

    private void populateInitialValues()
    {
      dataTable = new DataTable();
      folderBrowserDialog2 = new FolderBrowserDialog();
      dataTable.Columns.Add("Radius(r)");
      dataTable.Columns.Add("Tangential Stress");
      dataTable.Columns.Add("Longitudinal Stress");
      dataTable.Columns.Add("Radial Stress");
      dataGridView1.DataSource = dataTable;
      comboBox1.Items.Add("Yes");
      comboBox1.Items.Add("No");
      label10.Text = DateTime.Now.ToString();
      monthCalendar1.Visible = false;
      toolTip1.SetToolTip(this.label10, "Click to display the calendar");
      toolTip2.SetToolTip(this.button1, "Click to compute the results");
      toolTip1.SetToolTip(this.button2, "Click to reset the textboxes");
      SetStyle(ControlStyles.AllPaintingInWmPaint |
               ControlStyles.DoubleBuffer |
               ControlStyles.ResizeRedraw |
               ControlStyles.UserPaint,
               true);
      label21.Visible = false;
    }

    private void reset()
    {
      dataTable.Clear();
      comboBox1.Text = "Select";
      FormUtil.clearTextBoxes(Controls);
      resetGraphImage();
    }

    private void resetGraphImage()
    {
      if (pictureBox1.Image != null)
      {
        pictureBox1.Image.Dispose();
        pictureBox1.Image = null;
        pictureBox1.ResetText();
      }
    }
  }
}