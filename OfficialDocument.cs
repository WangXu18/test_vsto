using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AddInDesignerObjects;
using Office;
using System.Diagnostics;

namespace vsto
{
  public class OfficialDocument : IDTExtensibility2, IRibbonExtensibility {
    public static Word.Application app = null;
    public static object wps;
    public static Word.Document wordDoc;
    Object Nothing = System.Reflection.Missing.Value;
    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) {
      //wps = Application;
      //app = wps as Word.Application;

      //wordDoc = app.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
      ////wordDoc = app.ActiveDocument;
      //wordDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
      //wordDoc.PageSetup.TopMargin = app.CentimetersToPoints(3.7f); // 37mm,对应104.9磅(1磅约等于0.3572mm)
      //wordDoc.PageSetup.BottomMargin = app.CentimetersToPoints(3.5f);
      //wordDoc.PageSetup.LeftMargin = app.CentimetersToPoints(2.8f);
      //wordDoc.PageSetup.RightMargin = app.CentimetersToPoints(2.7f);
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) {
      throw new NotImplementedException();
    }

    public void OnAddInsUpdate(ref Array custom) {
      throw new NotImplementedException();
    }

    public void OnStartupComplete(ref Array custom) {
      throw new NotImplementedException();
    }

    public void OnBeginShutdown(ref Array custom) {
      throw new NotImplementedException();
    }

    string IRibbonExtensibility.GetCustomUI(string RibbonID) {
      return Properties.Resource1.MyRibbon;
    }

    public void reduceFontSpace(IRibbonControl ctrl) {
      MessageBox.Show("you click reduceFontSpace");
    }

    public Bitmap GetRibbonImage(IRibbonControl ctrl) {
      switch (ctrl.Id) {
        case "rh":
          return new Bitmap(Properties.Resource1.chrome);
        default:
          return new Bitmap(Properties.Resource1.design);
      }
    }

    //public Bitmap LoadImage(IRibbonControl ctrl) {
    //  switch (ctrl.Id)
    //  {
    //    case "rh":
    //      return new Bitmap(Properties.Resource1.chrome);
    //    default:
    //      return new Bitmap(Properties.Resource1.design);
    //  }
    //}

    public void setCommonRH2(IRibbonControl ctrl) {
      Process.Start("F:\\Code\\vs\\vsto\\exe\\ConvertBeikeLib.exe");
    }
  }
}
