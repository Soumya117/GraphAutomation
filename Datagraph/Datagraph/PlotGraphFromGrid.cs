using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace datagraph
{
  class PlotGraphFromGrid : IPlotGraph
  {
    private ExcelObject excelObject;
    private DataTable dataTable;
    private DataGridView dataGridView;

    public PlotGraphFromGrid(ExcelObject excel, DataTable data, DataGridView dataGrid)
    {
      excelObject = excel;
      dataTable = data;
      dataGridView = dataGrid;
    }

    void IPlotGraph.plot()
    {
      var path = excelObject.configurePath();

      string filename = @path + "\\exprt.xls";
      try
      {
        excelObject.configureExcelApp();
        var worksheet = excelObject.configureWorksheet();

        for (int i = 1; i < dataGridView.Columns.Count + 1; i++)
        {
          worksheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
        }

        for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
        {
          for (int j = 0; j < dataGridView.Columns.Count; j++)
          {
            worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
          }
        }
        excelObject.saveToTempExcel(filename);
        excelObject.configureChartsAndSave();
      }
      catch (System.Exception ex)
      {
        MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      finally
      {
        excelObject.releaseObjects();
      }
    }
  }
}
