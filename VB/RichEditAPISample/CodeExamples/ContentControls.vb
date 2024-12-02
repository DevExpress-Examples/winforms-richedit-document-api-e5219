Imports DevExpress.XtraRichEdit
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditAPISample.CodeExamples

    Friend Class ContentControls

        Private Shared Sub CreateContentControls(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#CreateContentControls"
            document.LoadDocument("Documents\Simple Form.docx")
            Dim contentControls = document.ContentControls
            ' Insert a form to enter a name:
            Dim namePosition = document.CreatePosition(document.Paragraphs(CInt((0))).Range.[End].ToInt() - 1)
            Dim nameControl = contentControls.InsertPlainTextControl(namePosition)
            ' Insert text in a content control:
            Dim nameTextPosition = document.CreatePosition(nameControl.Range.Start.ToInt() + 1)
            document.InsertText(nameTextPosition, "Click to enter a name")
            ' Insert a drop-down list to select the appointment type:
            Dim listPosition = document.CreatePosition(document.Paragraphs(CInt((1))).Range.[End].ToInt() - 1)
            Dim listControl = contentControls.InsertDropDownListControl(listPosition)
            ' Add items to the drop-down list:
            listControl.AddItem("First Appointment", "First Appointment")
            listControl.AddItem("Follow-Up Appointment", "Follow-Up Appointment")
            listControl.AddItem("Laboratory Results Check", "Laboratory Results Check")
            listControl.SelectedItemIndex = 1
            ' Insert a date picker to select the appointment date:
            Dim datePosition = document.CreatePosition(document.Paragraphs(CInt((2))).Range.[End].ToInt() - 1)
            Dim datePicker = contentControls.InsertDatePickerControl(datePosition)
            datePicker.DateFormat = "dddd, MMMM dd, yyyy"
            ' Insert a checkbox:
            Dim checkboxControl = contentControls.InsertCheckboxControl(document.Paragraphs(CInt((3))).Range.Start)
            checkboxControl.Checked = False
#End Region  ' #CreateContentControls
        End Sub

        Private Shared Sub ChangeContentControls(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#ChangeContentControlParameters"
            document.LoadDocument("Documents\Simple Form Filled.docx")
            Dim contentControls = document.ContentControls
            For Each contentControl In contentControls
                contentControl.Color = System.Drawing.Color.Red
                Select Case contentControl.ControlType
                    Case DevExpress.XtraRichEdit.API.Native.ContentControlType.RichText, DevExpress.XtraRichEdit.API.Native.ContentControlType.PlainText
                        contentControl.IsTemporary = True
                    Case DevExpress.XtraRichEdit.API.Native.ContentControlType.Checkbox
                        Dim checkbox As DevExpress.XtraRichEdit.API.Native.ContentControlCheckbox = TryCast(contentControl, DevExpress.XtraRichEdit.API.Native.ContentControlCheckbox)
                        checkbox.CheckedSymbolStyle.Character = "*"c
                End Select
            Next
#End Region  ' #ChangeContentControlParameters        
        End Sub

        Private Shared Sub RemoveContentControls(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#RemoveContentControls"
            document.LoadDocument("Documents\Simple Form Filled.docx")
            Dim contentControls = document.ContentControls
            For i = 0 To contentControls.Count - 1
                If contentControls(CInt((i))).ControlType = DevExpress.XtraRichEdit.API.Native.ContentControlType.[Date] Then
                    contentControls.Remove(contentControls(i), True)
                End If
            Next
#End Region  ' #RemoveContentControls
        End Sub
    End Class
End Namespace
