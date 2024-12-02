Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditAPISample.CodeExamples

    Friend Class PageLayoutActions

        Private Shared Sub LineNumbering(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#LineNumbering"
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            Dim sec As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            sec.LineNumbering.CountBy = 2
            sec.LineNumbering.Start = 1
            sec.LineNumbering.Distance = 0.25F
            sec.LineNumbering.RestartType = DevExpress.XtraRichEdit.API.Native.LineNumberingRestart.NewSection
#End Region  ' #LineNumbering
        End Sub

        Private Shared Sub CreateColumns(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#CreateColumns"
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Get the first section in a document.
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Create equal width column layout.
            Dim sectionColumnsLayout As DevExpress.XtraRichEdit.API.Native.SectionColumnCollection = firstSection.Columns.CreateUniformColumns(firstSection.Page, 0.2F, 3)
            ' Set different column width.
            sectionColumnsLayout(CInt((0))).Width = 3F
            sectionColumnsLayout(CInt((1))).Width = 2F
            sectionColumnsLayout(CInt((2))).Width = 1F
            ' Apply layout to the document.
            firstSection.Columns.SetColumns(sectionColumnsLayout)
#End Region  ' #CreateColumns
        End Sub

        Private Shared Sub PrintLayout(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#PrintLayout"
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            document.Sections(CInt((0))).Page.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A6
            document.Sections(CInt((0))).Page.Landscape = True
            document.Sections(CInt((0))).Margins.Left = 2.0F
#End Region  ' #PrintLayout
        End Sub

        Private Shared Sub TabStops(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#TabStops"
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            Dim tabs As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = document.Paragraphs(CInt((0))).BeginUpdateTabs(True)
            Dim tab1 As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            ' Sets tab stop at 2.5 inch.
            tab1.Position = 2.5F
            tab1.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.Left
            tab1.Leader = DevExpress.XtraRichEdit.API.Native.TabLeaderType.MiddleDots
            tabs.Add(tab1)
            Dim tab2 As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            tab2.Position = 5.5F
            tab2.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.[Decimal]
            tab2.Leader = DevExpress.XtraRichEdit.API.Native.TabLeaderType.EqualSign
            tabs.Add(tab2)
            document.Paragraphs(CInt((0))).EndUpdateTabs(tabs)
#End Region  ' #TabStops
        End Sub

        Private Shared Sub CreatePageBorders(ByVal document As DevExpress.XtraRichEdit.API.Native.Document)
#Region "#CreatePageBorders"
            ' Generate a document with two sections and multiple pages in each section.
            document.AppendText(Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed)
            document.Paragraphs.Append()
            document.AppendSection()
            document.AppendText(Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed)
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            Dim pageBorders1 As DevExpress.XtraRichEdit.API.Native.SectionPageBorders = firstSection.PageBorders
            ' Set page borders for the first page of the first section.
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders1.LeftBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders1.TopBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders1.RightBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders1.BottomBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            pageBorders1.AppliesTo = DevExpress.XtraRichEdit.API.Native.PageBorderAppliesTo.FirstPage
            Dim secondSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(1)
            Dim pageBorders2 As DevExpress.XtraRichEdit.API.Native.SectionPageBorders = secondSection.PageBorders
            ' Set page borders for all pages of the second section.
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders2.LeftBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders2.TopBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders2.RightBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            Call RichEditAPISample.CodeExamples.PageLayoutActions.PageBorderHelper.SetPageBorders(pageBorders2.BottomBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            pageBorders2.AppliesTo = DevExpress.XtraRichEdit.API.Native.PageBorderAppliesTo.AllPages
            pageBorders2.ZOrder = DevExpress.XtraRichEdit.API.Native.PageBorderZOrder.Back
#End Region  ' #CreatePageBorders
        End Sub

#Region "#@CreatePageBorders"
        Private Class PageBorderHelper

            Public Shared Sub SetPageBorders(ByVal border As DevExpress.XtraRichEdit.API.Native.PageBorder, ByVal lineStyle As DevExpress.XtraRichEdit.API.Native.BorderLineStyle, ByVal borderWidth As Single, ByVal color As System.Drawing.Color)
                border.LineStyle = lineStyle
                border.LineWidth = borderWidth
                border.LineColor = color
            End Sub
        End Class
#End Region  ' #@CreatePageBorders
    End Class
End Namespace
