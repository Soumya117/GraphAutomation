using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace datagraph
{
  public static class FormUtil
  {
    public static void ClearTextBoxes(Control.ControlCollection controls)
    {
      Action<Control.ControlCollection> func = null;
      func = (controls1) =>
      {
        foreach (Control control in controls1)
          if (control is TextBox)
            (control as TextBox).Clear();
          else
            func(control.Controls);
      };
      func(controls);
    }

    public static void verifyTextBoxInputIsNum(TextBox textBox)
    {
      float textValue;
      if (!String.IsNullOrEmpty(textBox.Text) && !float.TryParse(textBox.Text, out textValue))
      {
        MessageBox.Show("Only Numbers are allowed", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Error);
        textBox.ResetText();
      }
    }

    public static string setFilePath(string selectedPath, bool isDefaultLoc, bool isRoot)
    {
      string path;
      if (isDefaultLoc == false)
      {
        path = @selectedPath + "Export";
      }
      else if (isRoot == false)
      {
        path = @selectedPath + "\\Export";
      }
      else
      {
        path = "C:\\Export";
      }
      return path;
    }
  }
}
