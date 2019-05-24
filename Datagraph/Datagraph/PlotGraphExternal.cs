using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace datagraph
{
  class PlotGraphExternal : IPlotGraph
  {
    private ExcelObject excelObject;
    private DataTable dataTable;
    private String fileName;

    public PlotGraphExternal(ExcelObject excel, DataTable data, String file)
    {
      excelObject = excel;
      dataTable = data;
      fileName = file;
    }

    void IPlotGraph.plot()
    {
        try
        {
          excelObject.configurePath();
          excelObject.configureExcelApp();
          excelObject.exportExternalFile(fileName);
          excelObject.configureChartsAndSave();
        }
        catch (System.Exception ex)
        {
          MessageBox.Show(ex.ToString(), "Error during exporting data", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
          excelObject.releaseObjects();
        }
      }
    }
  }
