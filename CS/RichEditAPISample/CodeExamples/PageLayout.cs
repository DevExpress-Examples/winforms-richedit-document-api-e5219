using DevExpress.XtraRichEdit.API.Native;

namespace RichEditAPISample.CodeExamples
{
    class PageLayoutActions
    {
        static void LineNumbering(Document document)
        {
            #region #LineNumbering
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            Section sec = document.Sections[0];
            sec.LineNumbering.CountBy = 2;
            sec.LineNumbering.Start = 1;
            sec.LineNumbering.Distance = 0.25f;
            sec.LineNumbering.RestartType = LineNumberingRestart.NewSection;
            #endregion #LineNumbering
        }

        static void CreateColumns(Document document)
        {
            #region #CreateColumns
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            // Get the first section in a document.
            Section firstSection = document.Sections[0];
            // Create equal width column layout.
            SectionColumnCollection sectionColumnsLayout =
                firstSection.Columns.CreateUniformColumns(firstSection.Page, 0.2f, 3);
            // Set different column width.
            sectionColumnsLayout[0].Width = 3f;
            sectionColumnsLayout[1].Width = 2f;
            sectionColumnsLayout[2].Width = 1f;
            // Apply layout to the document.
            firstSection.Columns.SetColumns(sectionColumnsLayout);
            #endregion #CreateColumns
        }

        static void PrintLayout(Document document)
        {
            #region #PrintLayout
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            document.Sections[0].Page.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A6;
            document.Sections[0].Page.Landscape = true;
            document.Sections[0].Margins.Left = 2.0f;
            #endregion #PrintLayout
        }

        static void TabStops(Document document)
        {
            #region #TabStops
            document.Unit = DevExpress.Office.DocumentUnit.Inch;
            TabInfoCollection tabs = document.Paragraphs[0].BeginUpdateTabs(true);
            DevExpress.XtraRichEdit.API.Native.TabInfo tab1 = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            // Sets tab stop at 2.5 inch.
            tab1.Position = 2.5f;
            tab1.Alignment = TabAlignmentType.Left;
            tab1.Leader = TabLeaderType.MiddleDots;
            tabs.Add(tab1);
            DevExpress.XtraRichEdit.API.Native.TabInfo tab2 = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            tab2.Position = 5.5f;
            tab2.Alignment = TabAlignmentType.Decimal;
            tab2.Leader = TabLeaderType.EqualSign;
            tabs.Add(tab2);
            document.Paragraphs[0].EndUpdateTabs(tabs);
            #endregion #TabStops
        }
        static void CreatePageBorders(Document document)
        {
            #region #CreatePageBorders
            // Generate a document with two sections and multiple pages in each section.
            document.AppendText("\f\f\f");
            document.Paragraphs.Append();
            document.AppendSection();
            document.AppendText("\f\f");

            Section firstSection = document.Sections[0];
            SectionPageBorders pageBorders1 = firstSection.PageBorders;

            // Set page borders for the first page of the first section.
            PageBorderHelper.SetPageBorders(pageBorders1.LeftBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            PageBorderHelper.SetPageBorders(pageBorders1.TopBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            PageBorderHelper.SetPageBorders(pageBorders1.RightBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            PageBorderHelper.SetPageBorders(pageBorders1.BottomBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            pageBorders1.AppliesTo = PageBorderAppliesTo.FirstPage;

            Section secondSection = document.Sections[1];
            SectionPageBorders pageBorders2 = secondSection.PageBorders;

            // Set page borders for all pages of the second section.
            PageBorderHelper.SetPageBorders(pageBorders2.LeftBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            PageBorderHelper.SetPageBorders(pageBorders2.TopBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            PageBorderHelper.SetPageBorders(pageBorders2.RightBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            PageBorderHelper.SetPageBorders(pageBorders2.BottomBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            pageBorders2.AppliesTo = PageBorderAppliesTo.AllPages;
            pageBorders2.ZOrder = PageBorderZOrder.Back;
            #endregion #CreatePageBorders

        }
        #region #@CreatePageBorders
        class PageBorderHelper
        {
            public static void SetPageBorders(PageBorder border, BorderLineStyle lineStyle,
            float borderWidth, System.Drawing.Color color)
            {
                border.LineStyle = lineStyle;
                border.LineWidth = borderWidth;
                border.LineColor = color;
            }
        }
        #endregion #@CreatePageBorders
    }
}




