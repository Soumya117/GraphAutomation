using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Web;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

namespace datagraph
{
  struct ImageObject
  {
    public string imagenew;
    public string exclimage;
  }

  struct DirectoryFlags
  {
    public bool isDefaultLoc;
    public bool isRoot;
    public string selectedPath;
  }

  partial class Form1
  {
    class ExcelObject
    {
      private Excel.Application app = null;
      private Excel.Workbook workBook = null;
      private Excel.Worksheet workSheet = null;
      private Excel.Range chartRange = null;
      private Excel.ChartObjects charts = null;
      private Excel.ChartObject myChart = null;
      private Excel.Chart chartPage = null;
      private object misValue = System.Reflection.Missing.Value;
      private ImageObject imageObject;
      private Image imageFile;
      public DataTable datatable = null;
      private DirectoryFlags flags;

      public void setDirectoryFlags(
        bool isDefaultLoc, 
        bool isRoot,
        string selectedPath)
      {
        flags.isDefaultLoc = isDefaultLoc;
        flags.isRoot = isRoot;
        flags.selectedPath = selectedPath;
      }

      public void configureChartsAndSave()
      {
        configureCharts();
        saveWorkbook();
      }

      public Excel.Worksheet configureWorksheet()
      {
        workBook = app.Workbooks.Add(Type.Missing);
        workSheet = (Excel.Worksheet)workBook.Sheets["Sheet1"];
        workSheet = (Excel.Worksheet)workBook.ActiveSheet;
        return workSheet;
      }

      public Image getImage()
      {
        return imageFile;
      }

      public void exportExternalFile(string filename)
      {
        workBook = app.Workbooks.Open(filename, 0, true, 5, "", "", true,
                                       Excel.XlPlatform.xlWindows, "\t",
                                       false, false, 0, true, 1, 0);
        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
      }

      public void saveToTempExcel(string filename)
      {
        workBook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        workBook = app.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
      }

      public string configurePath()
      {
        string folderdate = DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss");
        string path = FormUtil.setFilePath(
          flags.selectedPath,
          flags.isDefaultLoc,
          flags.isRoot);
        path = path + "\\" + folderdate;

        if (!Directory.Exists(path))
        {
          System.IO.Directory.CreateDirectory(path);
        }

        imageObject.imagenew = @path + "\\Image.jpeg";
        imageObject.exclimage = @path + "\\exclimg.xls";
        return path;
      }

      public void configureExcelApp()
      {
        app = new Excel.Application();
        app.DisplayAlerts = false;
        app.Visible = false;
      }

      public void configureCharts()
      {
        charts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
        myChart = (Excel.ChartObject)charts.Add(10, 80, 300, 250);
        chartPage = myChart.Chart;
        myChart.Height = 700;
        myChart.Width = 1024;
        chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
        chartRange = workSheet.get_Range("A1", "D2000");
        chartPage.SetSourceData(chartRange, misValue);

        chartPage.Export(imageObject.imagenew, "JPEG", misValue);
        Image image = Image.FromFile(imageObject.imagenew);
        imageFile = image;
      }

      public void saveWorkbook()
      {
        workBook.SaveAs(imageObject.exclimage, Type.Missing,
        Type.Missing, Type.Missing,
        Type.Missing, Type.Missing,
        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
        Type.Missing, Type.Missing,
        Type.Missing, Type.Missing,
        Type.Missing);
        workBook.Close(true, misValue, misValue);
      }

      public void releaseObjects()
      {
        GarbageCollector.releaseObject(workSheet);
        GarbageCollector.releaseObject(workBook);
        GarbageCollector.releaseObject(app);
      }

      public void plotGraphFromDataGrid(DataTable dataTable, DataGridView dataGridView1)
      {
        var dataTableCount = System.Convert.ToInt32(dataTable.Rows.Count);
        if (dataTableCount == 0)
        {
          MessageBox.Show("No data points loaded");
          return;
        }

        var path = configurePath();

        string filename = @path + "\\exprt.xls";
        try
        {
          configureExcelApp();
          var worksheet = configureWorksheet();

          for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
          {
            worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
          }

          for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
          {
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
              worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            }
          }
          saveToTempExcel(filename);
          configureChartsAndSave();
        }
        catch (System.Exception ex)
        {
          MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
          releaseObjects();
        }
      }
      public void plotGraphFromExcel(DataTable dataTable, string fileName)
      {
        try
        {
          if (!String.IsNullOrEmpty(fileName))
          {
            dataTable.Clear();
            configurePath();
            configureExcelApp();
            exportExternalFile(fileName);
          }
          else
            return;

          configureChartsAndSave();
        }
        catch (System.Exception ex)
        {
          MessageBox.Show(ex.ToString(), "Error during exporting data", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
          releaseObjects();
        }
      }
    }
  }
}