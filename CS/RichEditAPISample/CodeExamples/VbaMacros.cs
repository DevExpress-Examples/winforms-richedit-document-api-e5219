using DevExpress.XtraRichEdit.API.Native;
using System.Windows.Forms;

namespace RichEditAPISample.CodeExamples
{
    class VbaMacrosActions
    {
        static void ObtainVbaModuleNames(Document document)
        {
            #region #ObtainVbaModuleNames
            document.LoadDocument("Documents\\Grimm.docx");
            if (document.VbaProject.Modules.Count > 0)
            {
                foreach (VbaModule module in document.VbaProject.Modules)
                { document.AppendText("\r\n \u00B7 " + module.Name); }
            }
            else
            {
                MessageBox.Show("This document does not contain any VBA modules");
            }
            #endregion #ObtainVbaModuleNames
        }

        static void ClearVbaModules(Document document)
        {
            #region #ClearVbaModules
            document.LoadDocument("Documents\\Grimm.docx");
            if (document.VbaProject.Modules.Count > 0)
                document.VbaProject.Modules.Clear();
            #endregion #ClearVbaModules
        }
    }
}
