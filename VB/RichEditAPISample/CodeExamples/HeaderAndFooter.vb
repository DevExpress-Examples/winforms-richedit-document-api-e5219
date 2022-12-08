Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace RichEditAPISample.CodeExamples

    Friend Class HeaderAndFooterActions

        Private Shared Sub ModifyHeader(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
'#Region "#ModifyHeader"
            document.AppendSection()
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Modify the header of the HeaderFooterType.First type.
            Dim myHeader As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.First)
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ")
            Dim fld As DevExpress.XtraRichEdit.API.Native.Field = myHeader.Fields.Create(range.[End], "PAGE \* ARABICDASH")
            myHeader.Fields.Update()
            firstSection.EndUpdateHeader(myHeader)
            ' Display the header of the HeaderFooterType.First type on the first page.
            firstSection.DifferentFirstPage = True
'#End Region  ' #ModifyHeader
        End Sub
    End Class
End Namespace
