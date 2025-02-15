﻿using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Drawing;
using System.IO;

namespace RichEditAPISample.CodeExamples
{
    class TableActions
    {
        static void CreateTable(Document document)
        {
            #region #CreateTable
            // Insert new table.
            Table tbl = document.Tables.Create(document.Range.Start, 1, 3, AutoFitBehaviorType.AutoFitToWindow);
            // Create a table header.
            document.InsertText(tbl[0, 0].Range.Start, "Name");
            document.InsertText(tbl[0, 1].Range.Start, "Size");
            document.InsertText(tbl[0, 2].Range.Start, "DateTime");
            // Insert table data.
            DirectoryInfo dirinfo = new DirectoryInfo("C:\\");
            try
            {
                tbl.BeginUpdate();
                foreach (FileInfo fi in dirinfo.GetFiles())
                {
                    TableRow row = tbl.Rows.Append();
                    TableCell cell = row.FirstCell;
                    string fileName = fi.Name;
                    string fileLength = String.Format("{0:N0}", fi.Length);
                    string fileLastTime = String.Format("{0:g}", fi.LastWriteTime);
                    document.InsertSingleLineText(cell.Range.Start, fileName);
                    document.InsertSingleLineText(cell.Next.Range.Start, fileLength);
                    document.InsertSingleLineText(cell.Next.Next.Range.Start, fileLastTime);
                }
                // Center the table header.
                foreach (Paragraph p in document.Paragraphs.Get(tbl.FirstRow.Range))
                {
                    p.Alignment = ParagraphAlignment.Center;
                }
            }
            finally
            {
                tbl.EndUpdate();
            }
            #endregion #CreateTable
        }

        static void CreateFixedTable(Document document)
        {
            #region #CreateFixedTable
            Table table = document.Tables.Create(document.Range.Start, 3, 4);

            table.TableAlignment = TableRowAlignment.Center;
            table.TableLayout = TableLayoutType.Fixed;
            table.PreferredWidthType = WidthType.Fixed;
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4f);

            table.Rows[1].HeightType = HeightType.Exact;
            table.Rows[1].Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8f);

            table[1, 2].PreferredWidthType = WidthType.Fixed;
            table[1, 2].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);

            #endregion #CreateFixedTable
        }
        static void ChangeTableColor(Document document)
        {
            #region #ChangeTableColor
            // Create a table.
            Table table = document.Tables.Create(document.Range.Start, 3, 5, AutoFitBehaviorType.AutoFitToWindow);
            table.BeginUpdate();
            // Provide the space between table cells.
            // The distance between cells will be 4 mm.
            document.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            table.TableCellSpacing = 2;
            // Change the color of empty space between cells.
            table.TableBackgroundColor = Color.Violet;
            //Change cell background color.
            table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellColor));
            table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellBorderColor));
            table.EndUpdate();
            #endregion #ChangeTableColor

        }
        #region #@ChangeTableColor
        class TableHelper
        {
            public static void ChangeCellColor(TableCell cell, int i, int j)
            {
                cell.BackgroundColor = System.Drawing.Color.Yellow;
            }

            public static void ChangeCellBorderColor(TableCell cell, int i, int j)
            {
                cell.Borders.Bottom.LineColor = System.Drawing.Color.Red;
                cell.Borders.Left.LineColor = System.Drawing.Color.Red;
                cell.Borders.Right.LineColor = System.Drawing.Color.Red;
                cell.Borders.Top.LineColor = System.Drawing.Color.Red;
            }
        }
        #endregion #@ChangeTableColor
        static void CreateAndApplyTableStyle(Document document)
        {
            #region #CreateAndApplyTableStyle
            document.BeginUpdate();
            // Create a new table style.
            TableStyle tStyleMain = document.TableStyles.CreateNew();
            // Specify style characteristics.
            tStyleMain.AllCaps = true;
            tStyleMain.FontName = "Segoe Condensed";
            tStyleMain.FontSize = 14;
            tStyleMain.Alignment = ParagraphAlignment.Center;
            tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = BorderLineStyle.Dotted;
            tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = BorderLineStyle.Dotted;
            tStyleMain.TableBorders.Top.LineThickness = 1.5f;
            tStyleMain.TableBorders.Top.LineStyle = BorderLineStyle.Double;
            tStyleMain.TableBorders.Left.LineThickness = 1.5f;
            tStyleMain.TableBorders.Left.LineStyle = BorderLineStyle.Double;
            tStyleMain.TableBorders.Bottom.LineThickness = 1.5f;
            tStyleMain.TableBorders.Bottom.LineStyle = BorderLineStyle.Double;
            tStyleMain.TableBorders.Right.LineThickness = 1.5f;
            tStyleMain.TableBorders.Right.LineStyle = BorderLineStyle.Double;
            tStyleMain.CellBackgroundColor = System.Drawing.Color.LightBlue;
            tStyleMain.TableLayout = TableLayoutType.Fixed;
            tStyleMain.Name = "MyTableStyle";
            //Add the style to the document.
            document.TableStyles.Add(tStyleMain);
            document.EndUpdate();
            document.BeginUpdate();
            // Create a table.
            Table table = document.Tables.Create(document.Range.Start, 3, 3);
            table.TableLayout = TableLayoutType.Fixed;
            table.PreferredWidthType = WidthType.Fixed;
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4.5f);
            table[1, 1].PreferredWidthType = WidthType.Fixed;
            table[1, 1].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            // Apply a previously defined style.
            table.Style = tStyleMain;
            document.EndUpdate();

            document.InsertText(table[1, 1].Range.Start, "STYLED");
            #endregion #CreateAndApplyTableStyle
        }

        static void UseConditionalStyle(Document document)
        {
            #region #UseConditionalStyle
            document.LoadDocument("Documents//TableStyles.docx", DocumentFormat.OpenXml);
            document.BeginUpdate();

            // Create a new style that is based on the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
            TableStyle myNewStyle = document.TableStyles.CreateNew();
            myNewStyle.Parent = document.TableStyles["Grid Table 5 Dark Accent 1"];
            // Create conditional styles (styles for table elements)
            TableConditionalStyle myNewStyleForFirstRow =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow);
            myNewStyleForFirstRow.CellBackgroundColor = Color.PaleVioletRed;
            TableConditionalStyle myNewStyleForFirstColumn =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn);
            myNewStyleForFirstColumn.CellBackgroundColor = Color.PaleVioletRed;
            TableConditionalStyle myNewStyleForOddColumns =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding);
            myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(Color.PaleVioletRed);
            TableConditionalStyle myNewStyleForEvenColumns =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding);
            myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(Color.PaleVioletRed);
            document.TableStyles.Add(myNewStyle);
            // Create a new table and apply a new style.
            Table table = document.Tables.Create(document.Range.End, 4, 4, AutoFitBehaviorType.AutoFitToWindow);
            table.Style = myNewStyle;
            // Specify which conditional styles are in effect.
            table.TableLook = TableLookTypes.ApplyFirstRow | TableLookTypes.ApplyFirstColumn;

            document.EndUpdate();
            #endregion #UseConditionalStyle
        }

        static void ChangeColumnAppearance(Document document)
        {
            #region #ChangeColumnAppearance
            Table table = document.Tables.Create(document.Range.Start, 3, 10);
            table.BeginUpdate();
            //Change cell background color and vertical alignment in the third column.
            table.ForEachRow(new TableRowProcessorDelegate(ChangeColumnAppearanceHelper.ChangeColumnColor));
            table.EndUpdate();
            #endregion #ChangeColumnAppearance

        }
        #region #@ChangeColumnAppearance
        class ChangeColumnAppearanceHelper
        {
            public static void ChangeColumnColor(TableRow row, int rowIndex)
            {
                row[2].BackgroundColor = System.Drawing.Color.LightCyan;
                row[2].VerticalAlignment = TableCellVerticalAlignment.Center;
            }
        }
        #endregion #@ChangeColumnAppearance

        static void UseTableCellProcessor(Document document)
        {
            #region #UseTableCellProcessor
            Table table = document.Tables.Create(document.Range.Start, 8, 8);
            table.BeginUpdate();
            table.ForEachCell(new TableCellProcessorDelegate(UseTableCellProcessorHelper.MakeMultiplicationCell));
            table.EndUpdate();
            #endregion #UseTableCellProcessor
        }
        #region #@UseTableCellProcessor
        class UseTableCellProcessorHelper
        {
            public static void MakeMultiplicationCell(TableCell cell, int i, int j)
            {
                SubDocument doc = cell.Range.BeginUpdateDocument();
                doc.InsertText(cell.Range.Start,
                    String.Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)));
                cell.Range.EndUpdateDocument(doc);
            }
        }
        #endregion #@UseTableCellProcessor

        static void MergeCells(Document document)
        {
            #region #MergeCells
            Table table = document.Tables.Create(document.Range.Start, 6, 8);
            table.BeginUpdate();
            table.MergeCells(table[2, 1], table[5, 1]);
            table.MergeCells(table[2, 3], table[2, 7]);
            table.EndUpdate();
            #endregion #MergeCells
        }
        static void SplitCells(Document document)
        {
            #region #SplitCells
            Table table = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.FixedColumnWidth, 350);
            //split a cell into three: 
            table.Cell(2, 1).Split(1, 3);
            #endregion #SplitCells
        }
        static void DeleteTableElements(Document document)
        {
            #region #DeleteTableElements
            Table tbl = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToWindow);
            tbl.BeginUpdate();
            //Delete a cell:
            tbl.Cell(1, 1).Delete();
            //Delete a row:
            tbl.Rows[2].Delete();
            tbl.EndUpdate();
            #endregion #DeleteTableElements
        }
        static void DeleteTable(Document document)
        {
            #region #DeleteTable
            Table tbl = document.Tables.Create(document.Range.Start, 3, 4);
            //To delete the table, uncomment the method below:
            //  document.Tables.Remove(tbl);
            #endregion #DeleteTable 
        }
        static void WrapTextAroundTable(Document document)
        {
            #region #WrapTextAroundTable
            document.LoadDocument("Documents//Grimm.docx");

            Table table = document.Tables.Create(document.Paragraphs[4].Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToContents);

            table.BeginUpdate();
            table.TextWrappingType = TableTextWrappingType.Around;

            //Specify vertical alignment:
            table.RelativeVerticalPosition = TableRelativeVerticalPosition.Paragraph;
            table.VerticalAlignment = TableVerticalAlignment.None;
            table.OffsetYRelative = DevExpress.Office.Utils.Units.InchesToDocumentsF(2f);

            //Specify horizontal alignment:
            table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Margin;
            table.HorizontalAlignment = TableHorizontalAlignment.Center;

            //Set distance between the text and the table:
            table.MarginBottom = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginLeft = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginTop = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginRight = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.EndUpdate();
            #endregion #WrapTextAroundTable
        }
    }
}
