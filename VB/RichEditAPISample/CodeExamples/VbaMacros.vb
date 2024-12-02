Imports DevExpress.XtraRichEdit.API.Native
Imports System.Windows.Forms

Namespace RichEditAPISample.CodeExamples

    Friend Class VbaMacrosActions

        Private Shared Sub ObtainVbaModuleNames(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#ObtainVbaModuleNames"
            document.LoadDocument("Documents\Grimm.docx")
            If document.VbaProject.Modules.Count > 0 Then
                For Each [module] As DevExpress.XtraRichEdit.API.Native.VbaModule In document.VbaProject.Modules
                    document.AppendText(Global.Microsoft.VisualBasic.Constants.vbCrLf & " · " & [module].Name)
                Next
            Else
                Call System.Windows.Forms.MessageBox.Show("This document does not contain any VBA modules")
            End If
#End Region  ' #ObtainVbaModuleNames
        End Sub

        Private Shared Sub ClearVbaModules(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#ClearVbaModules"
            document.LoadDocument("Documents\Grimm.docx")
            If document.VbaProject.Modules.Count > 0 Then document.VbaProject.Modules.Clear()
#End Region  ' #ClearVbaModules
        End Sub
    End Class
End Namespace
