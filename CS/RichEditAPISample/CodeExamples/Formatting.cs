using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Drawing;

namespace RichEditAPISample.CodeExamples
{
    public static class FormattingActions
    {
        static void FormatText(Document document)
        {
            #region #FormatText
            document.BeginUpdate();
            document.AppendText("Normal\nFormatted\nNormal");
            document.EndUpdate();
            // The target range is the second paragraph 
            DocumentRange range = document.Paragraphs[1].Range;

            // Create and customize an object  
            // that sets character formatting for the selected range
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.FontName = "Comic Sans MS";
            cp.FontSize = 18;
            cp.ForeColor = Color.Blue;
            cp.BackColor = Color.Snow;
            cp.Underline = UnderlineType.DoubleWave;
            cp.UnderlineColor = Color.Red;

            // Finalize modifications  
            // with this method call 
            document.EndUpdateCharacters(cp);
            #endregion #FormatText
        }

        static void ResetCharacterFormatting(Document document)
        {
            #region #ResetCharacterFormatting
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            // Set font size and font name of the characters in the first paragraph to default. 
            // Other character properties remain intact.
            DocumentRange range = document.Paragraphs[0].Range;
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.Reset(CharacterPropertiesMask.FontSize | CharacterPropertiesMask.FontName);
            document.EndUpdateCharacters(cp);
            #endregion #ResetCharacterFormatting
        }

        static void FormatParagraph(Document document)
        {
            #region #FormatParagraph
            document.BeginUpdate();
            document.AppendText("Modified Paragraph\nNormal\nNormal");
            document.EndUpdate();

            //The target range is the first paragraph
            DocumentPosition pos = document.Range.Start;
            DocumentRange range = document.CreateRange(pos, 0);

            // Create and customize an object  
            // that sets character formatting for the selected range
            ParagraphProperties pp = document.BeginUpdateParagraphs(range);
            // Center paragraph
            pp.Alignment = ParagraphAlignment.Center;
            // Set triple spacing
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 3;
            // Set left indent at 0.5".
            // Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5f);
            // Set tab stop at 1.5"
            TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
            TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            tbi.Alignment = TabAlignmentType.Center;
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            tbiColl.Add(tbi);
            pp.EndUpdateTabs(tbiColl);

            //Finalize modifications
            // with this method call
            document.EndUpdateParagraphs(pp);
            #endregion #FormatParagraph
        }

        static void ResetParagraphFormatting(Document document)
        {
            #region #ResetParagraphFormatting
            document.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            // Set alignment and indentation of the first line in the first paragraph to default. 
            // Other paragraph properties remain intact.
            DocumentRange range = document.Paragraphs[0].Range;
            ParagraphProperties cp = document.BeginUpdateParagraphs(range);
            cp.Reset(ParagraphPropertiesMask.Alignment | ParagraphPropertiesMask.FirstLineIndent);
            document.EndUpdateParagraphs(cp);
            #endregion #ResetParagraphFormatting
        }

        static void FormatParagraphBorders(Document document)
        {
            #region #FormatParagraphBorders
            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText(String.Format("Modified Paragraph" +
                Environment.NewLine + "Normal" + Environment.NewLine + "Normal"));

            // Finalize to edit the document.
            document.EndUpdate();

            // Obtain the first and last paragraph ranges
            Paragraph firstParagraph = document.Paragraphs[0];
            Paragraph thirdParagraph = document.Paragraphs[2];
            DocumentRange paragraphRange = document.CreateRange(firstParagraph.Range.Start,
                            thirdParagraph.Range.End.ToInt() - firstParagraph.Range.Start.ToInt());

            // Start to edit the paragraph.
            ParagraphProperties pp = document.BeginUpdateParagraphs(paragraphRange);
            BorderHelper.SetBorder(pp.Borders.HorizontalBorder);
            BorderHelper.SetBorder(pp.Borders.BottomBorder);
            BorderHelper.SetBorder(pp.Borders.TopBorder);
            BorderHelper.SetBorder(pp.Borders.LeftBorder);
            BorderHelper.SetBorder(pp.Borders.RightBorder);

            // Finalize to edit the paragraph.
            document.EndUpdateParagraphs(pp);
            #endregion #FormatParagraphBorders
        }
        #region #@FormatParagraphBorders
        class BorderHelper
        {
            public static void SetBorder(ParagraphBorder border)
            {
                border.LineWidth = 2f;
                border.LineStyle = BorderLineStyle.Thick;
                border.LineColor = Color.SteelBlue;
            }
        }
        #endregion #@FormatParagraphBorders


        static void ResetParagraphFormatting(RichEditDocumentServer wordProcessor)
        {
            #region #ResetParagraphFormatting
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the range of the document's first paragraph.
            DocumentRange range = document.Paragraphs[0].Range;

            // Start to edit the paragraph.
            ParagraphProperties cp = document.BeginUpdateParagraphs(range);

            // Set alignmment and first line indent of the target paragraph to default values.   
            // Other paragraph properties remain intact.
            cp.Reset(ParagraphPropertiesMask.Alignment | ParagraphPropertiesMask.FirstLineIndent);

            // Finalize to edit the paragraph.
            document.EndUpdateParagraphs(cp);
            #endregion #ResetParagraphFormatting
        }




    }
}

