Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Drawing

Namespace RichEditAPISample.CodeExamples

    Public Module FormattingActions

        Private Sub FormatText(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#FormatText"
            document.BeginUpdate()
            document.AppendText("Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Formatted" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            document.EndUpdate()
            ' The target range is the second paragraph 
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((1))).Range
            ' Create and customize an object  
            ' that sets character formatting for the selected range
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            cp.FontName = "Comic Sans MS"
            cp.FontSize = 18
            cp.ForeColor = System.Drawing.Color.Blue
            cp.BackColor = System.Drawing.Color.Snow
            cp.Underline = DevExpress.XtraRichEdit.API.Native.UnderlineType.DoubleWave
            cp.UnderlineColor = System.Drawing.Color.Red
            ' Finalize modifications  
            ' with this method call 
            document.EndUpdateCharacters(cp)
#End Region  ' #FormatText
        End Sub

        Private Sub ResetCharacterFormatting(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#ResetCharacterFormatting"
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Set font size and font name of the characters in the first paragraph to default. 
            ' Other character properties remain intact.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            cp.Reset(DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontSize Or DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontName)
            document.EndUpdateCharacters(cp)
#End Region  ' #ResetCharacterFormatting
        End Sub

        Private Sub FormatParagraph(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#FormatParagraph"
            document.BeginUpdate()
            document.AppendText("Modified Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            document.EndUpdate()
            'The target range is the first paragraph
            Dim pos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Range.Start
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(pos, 0)
            ' Create and customize an object  
            ' that sets character formatting for the selected range
            Dim pp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Center paragraph
            pp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
            ' Set triple spacing
            pp.LineSpacingType = DevExpress.XtraRichEdit.API.Native.ParagraphLineSpacing.Multiple
            pp.LineSpacingMultiplier = 3
            ' Set left indent at 0.5".
            ' Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5F)
            ' Set tab stop at 1.5"
            Dim tbiColl As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = pp.BeginUpdateTabs(True)
            Dim tbi As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            tbi.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.Center
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            tbiColl.Add(tbi)
            pp.EndUpdateTabs(tbiColl)
            'Finalize modifications
            ' with this method call
            document.EndUpdateParagraphs(pp)
#End Region  ' #FormatParagraph
        End Sub

        Private Sub ResetParagraphFormatting(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#ResetParagraphFormatting"
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Set alignment and indentation of the first line in the first paragraph to default. 
            ' Other paragraph properties remain intact.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            cp.Reset(DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.Alignment Or DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.FirstLineIndent)
            document.EndUpdateParagraphs(cp)
#End Region  ' #ResetParagraphFormatting
        End Sub

        Private Sub FormatParagraphBorders(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#FormatParagraphBorders"
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText(System.[String].Format("Modified Paragraph" & System.Environment.NewLine & "Normal" & System.Environment.NewLine & "Normal"))
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Obtain the first and last paragraph ranges
            Dim firstParagraph As DevExpress.XtraRichEdit.API.Native.Paragraph = document.Paragraphs(0)
            Dim thirdParagraph As DevExpress.XtraRichEdit.API.Native.Paragraph = document.Paragraphs(2)
            Dim paragraphRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(firstParagraph.Range.Start, thirdParagraph.Range.[End].ToInt() - firstParagraph.Range.Start.ToInt())
            ' Start to edit the paragraph.
            Dim pp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(paragraphRange)
            Call RichEditAPISample.CodeExamples.FormattingActions.BorderHelper.SetBorder(pp.Borders.HorizontalBorder)
            Call RichEditAPISample.CodeExamples.FormattingActions.BorderHelper.SetBorder(pp.Borders.BottomBorder)
            Call RichEditAPISample.CodeExamples.FormattingActions.BorderHelper.SetBorder(pp.Borders.TopBorder)
            Call RichEditAPISample.CodeExamples.FormattingActions.BorderHelper.SetBorder(pp.Borders.LeftBorder)
            Call RichEditAPISample.CodeExamples.FormattingActions.BorderHelper.SetBorder(pp.Borders.RightBorder)
            ' Finalize to edit the paragraph.
            document.EndUpdateParagraphs(pp)
#End Region  ' #FormatParagraphBorders
        End Sub

#Region "#@FormatParagraphBorders"
        Private Class BorderHelper

            Public Shared Sub SetBorder(ByVal border As DevExpress.XtraRichEdit.API.Native.ParagraphBorder)
                border.LineWidth = 2F
                border.LineStyle = DevExpress.XtraRichEdit.API.Native.BorderLineStyle.Thick
                border.LineColor = System.Drawing.Color.SteelBlue
            End Sub
        End Class

#End Region  ' #@FormatParagraphBorders
        Private Sub ResetParagraphFormatting(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ResetParagraphFormatting"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the range of the document's first paragraph.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            ' Start to edit the paragraph.
            Dim cp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Set alignmment and first line indent of the target paragraph to default values.   
            ' Other paragraph properties remain intact.
            cp.Reset(DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.Alignment Or DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.FirstLineIndent)
            ' Finalize to edit the paragraph.
            document.EndUpdateParagraphs(cp)
#End Region  ' #ResetParagraphFormatting
        End Sub
    End Module
End Namespace
