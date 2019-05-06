using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;

namespace datagraph
{
  public static class GarbageCollector
  {
    public static void releaseObject(object obj)
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
  }
}
