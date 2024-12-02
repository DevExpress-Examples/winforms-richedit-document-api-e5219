Imports DevExpress.XtraTab
Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing
Imports DevExpress.Office.Utils

Namespace RichEditAPISample

    Public Partial Class Form1
        Inherits XtraForm

#Region "Controls"
        Private treeList1 As TreeList

        Private xtraTabControl1 As XtraTabControl

        Private xtraTabPage1 As XtraTabPage

        Private richEditControlCS As RichEditControl

        Private xtraTabPage2 As XtraTabPage

        Public displayResultControl1 As DisplayResultControl

        Private layoutControl1 As DevExpress.XtraLayout.LayoutControl

        Private Root As DevExpress.XtraLayout.LayoutControlGroup

        Private layoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem

        Private layoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem

        Private layoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem

        Private layoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem

        Private splitterItem2 As DevExpress.XtraLayout.SplitterItem

        Private splitterItem1 As DevExpress.XtraLayout.SplitterItem

        Private layoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup

        Private codeExampleNameLbl As DevExpress.XtraLayout.SimpleLabelItem

        Private richEditControlVB As RichEditControl

#End Region
#Region "InitializeComponent"
        Private Sub InitializeComponent()
            checkEdit1 = New CheckEdit()
            layoutControl1 = New DevExpress.XtraLayout.LayoutControl()
            treeList1 = New TreeList()
            displayResultControl1 = New DisplayResultControl()
            xtraTabControl1 = New XtraTabControl()
            xtraTabPage1 = New XtraTabPage()
            richEditControlCS = New RichEditControl()
            xtraTabPage2 = New XtraTabPage()
            richEditControlVB = New RichEditControl()
            xtraTabPage3 = New XtraTabPage()
            richEditControlCSClass = New RichEditControl()
            xtraTabPage4 = New XtraTabPage()
            richEditControlVBClass = New RichEditControl()
            Root = New DevExpress.XtraLayout.LayoutControlGroup()
            layoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem()
            splitterItem1 = New DevExpress.XtraLayout.SplitterItem()
            layoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup()
            splitterItem2 = New DevExpress.XtraLayout.SplitterItem()
            layoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem()
            layoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
            layoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
            codeExampleNameLbl = New DevExpress.XtraLayout.SimpleLabelItem()
            CType(checkEdit1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            layoutControl1.SuspendLayout()
            CType(treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            xtraTabControl1.SuspendLayout()
            xtraTabPage1.SuspendLayout()
            xtraTabPage2.SuspendLayout()
            xtraTabPage3.SuspendLayout()
            xtraTabPage4.SuspendLayout()
            CType(Root, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(splitterItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(splitterItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(layoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(codeExampleNameLbl, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            ' 
            ' checkEdit1
            ' 
            checkEdit1.AutoSizeInLayoutControl = True
            checkEdit1.Location = New System.Drawing.Point(596, 18)
            checkEdit1.Name = "checkEdit1"
            checkEdit1.Properties.Caption = "Indicate cursor position at window caption"
            checkEdit1.Size = New System.Drawing.Size(225, 20)
            checkEdit1.StyleController = layoutControl1
            checkEdit1.TabIndex = 12
            AddHandler checkEdit1.CheckedChanged, New EventHandler(AddressOf checkEdit1_CheckedChanged)
            ' 
            ' layoutControl1
            ' 
            layoutControl1.Controls.Add(treeList1)
            layoutControl1.Controls.Add(displayResultControl1)
            layoutControl1.Controls.Add(xtraTabControl1)
            layoutControl1.Controls.Add(checkEdit1)
            layoutControl1.Dock = DockStyle.Fill
            layoutControl1.Location = New System.Drawing.Point(0, 0)
            layoutControl1.Name = "layoutControl1"
            layoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(742, 351, 650, 403)
            layoutControl1.Root = Root
            layoutControl1.Size = New System.Drawing.Size(1248, 668)
            layoutControl1.TabIndex = 1
            layoutControl1.Text = "layoutControl1"
            ' 
            ' treeList1
            ' 
            treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline)
            treeList1.Appearance.FocusedCell.Options.UseFont = True
            treeList1.Location = New System.Drawing.Point(835, 12)
            treeList1.Name = "treeList1"
            treeList1.Size = New System.Drawing.Size(401, 644)
            treeList1.TabIndex = 11
            ' 
            ' displayResultControl1
            ' 
            displayResultControl1.Location = New System.Drawing.Point(12, 317)
            displayResultControl1.Name = "displayResultControl1"
            displayResultControl1.ReviewingPaneFormVisible = False
            displayResultControl1.Size = New System.Drawing.Size(809, 339)
            displayResultControl1.TabIndex = 0
            ' 
            ' xtraTabControl1
            ' 
            xtraTabControl1.AppearancePage.PageClient.BackColor = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.BackColor2 = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.BorderColor = System.Drawing.Color.Transparent
            xtraTabControl1.AppearancePage.PageClient.Options.UseBackColor = True
            xtraTabControl1.AppearancePage.PageClient.Options.UseBorderColor = True
            xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.True
            xtraTabControl1.Location = New System.Drawing.Point(12, 48)
            xtraTabControl1.Name = "xtraTabControl1"
            xtraTabControl1.SelectedTabPage = xtraTabPage1
            xtraTabControl1.Size = New System.Drawing.Size(809, 255)
            xtraTabControl1.TabIndex = 11
            xtraTabControl1.TabPages.AddRange(New XtraTabPage() {xtraTabPage1, xtraTabPage2, xtraTabPage3, xtraTabPage4})
            ' 
            ' xtraTabPage1
            ' 
            xtraTabPage1.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            xtraTabPage1.Appearance.HeaderActive.Options.UseFont = True
            xtraTabPage1.Controls.Add(richEditControlCS)
            xtraTabPage1.Name = "xtraTabPage1"
            xtraTabPage1.Size = New System.Drawing.Size(807, 230)
            xtraTabPage1.Tag = "CS"
            xtraTabPage1.Text = "CS"
            ' 
            ' richEditControlCS
            ' 
            richEditControlCS.ActiveViewType = RichEditViewType.Draft
            richEditControlCS.Dock = DockStyle.Fill
            richEditControlCS.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            richEditControlCS.Location = New System.Drawing.Point(0, 0)
            richEditControlCS.Name = "richEditControlCS"
            richEditControlCS.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlCS.Size = New System.Drawing.Size(807, 230)
            richEditControlCS.TabIndex = 14
            ' 
            ' xtraTabPage2
            ' 
            xtraTabPage2.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            xtraTabPage2.Appearance.HeaderActive.Options.UseFont = True
            xtraTabPage2.Controls.Add(richEditControlVB)
            xtraTabPage2.Name = "xtraTabPage2"
            xtraTabPage2.Size = New System.Drawing.Size(778, 181)
            xtraTabPage2.Tag = "VB"
            xtraTabPage2.Text = "VB"
            ' 
            ' richEditControlVB
            ' 
            richEditControlVB.ActiveViewType = RichEditViewType.Draft
            richEditControlVB.Dock = DockStyle.Fill
            richEditControlVB.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            richEditControlVB.Location = New System.Drawing.Point(0, 0)
            richEditControlVB.Name = "richEditControlVB"
            richEditControlVB.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlVB.Size = New System.Drawing.Size(778, 181)
            richEditControlVB.TabIndex = 15
            ' 
            ' xtraTabPage3
            ' 
            xtraTabPage3.Controls.Add(richEditControlCSClass)
            xtraTabPage3.Name = "xtraTabPage3"
            xtraTabPage3.Size = New System.Drawing.Size(778, 181)
            xtraTabPage3.Tag = "CS"
            xtraTabPage3.Text = "СS Helper"
            ' 
            ' richEditControlCSClass
            ' 
            richEditControlCSClass.ActiveViewType = RichEditViewType.Draft
            richEditControlCSClass.Dock = DockStyle.Fill
            richEditControlCSClass.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            richEditControlCSClass.Location = New System.Drawing.Point(0, 0)
            richEditControlCSClass.Name = "richEditControlCSClass"
            richEditControlCSClass.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlCSClass.Size = New System.Drawing.Size(778, 181)
            richEditControlCSClass.TabIndex = 0
            ' 
            ' xtraTabPage4
            ' 
            xtraTabPage4.Controls.Add(richEditControlVBClass)
            xtraTabPage4.Name = "xtraTabPage4"
            xtraTabPage4.Size = New System.Drawing.Size(778, 181)
            xtraTabPage4.Tag = "VB"
            xtraTabPage4.Text = "VB Helper"
            ' 
            ' richEditControlVBClass
            ' 
            richEditControlVBClass.ActiveViewType = RichEditViewType.Draft
            richEditControlVBClass.Dock = DockStyle.Fill
            richEditControlVBClass.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            richEditControlVBClass.Location = New System.Drawing.Point(0, 0)
            richEditControlVBClass.Name = "richEditControlVBClass"
            richEditControlVBClass.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden
            richEditControlVBClass.Size = New System.Drawing.Size(778, 181)
            richEditControlVBClass.TabIndex = 1
            ' 
            ' Root
            ' 
            Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True
            Root.GroupBordersVisible = False
            Root.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {layoutControlItem5, splitterItem1, layoutControlGroup1})
            Root.Name = "Root"
            Root.Size = New System.Drawing.Size(1248, 668)
            Root.TextVisible = False
            ' 
            ' layoutControlItem5
            ' 
            layoutControlItem5.Control = treeList1
            layoutControlItem5.Location = New System.Drawing.Point(823, 0)
            layoutControlItem5.Name = "layoutControlItem5"
            layoutControlItem5.Size = New System.Drawing.Size(405, 648)
            layoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
            layoutControlItem5.TextVisible = False
            ' 
            ' splitterItem1
            ' 
            splitterItem1.AllowHotTrack = True
            splitterItem1.Location = New System.Drawing.Point(813, 0)
            splitterItem1.Name = "splitterItem1"
            splitterItem1.Size = New System.Drawing.Size(10, 648)
            ' 
            ' layoutControlGroup1
            ' 
            layoutControlGroup1.GroupBordersVisible = False
            layoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {splitterItem2, layoutControlItem4, layoutControlItem2, layoutControlItem3, codeExampleNameLbl})
            layoutControlGroup1.Location = New System.Drawing.Point(0, 0)
            layoutControlGroup1.Name = "layoutControlGroup1"
            layoutControlGroup1.Size = New System.Drawing.Size(813, 648)
            ' 
            ' splitterItem2
            ' 
            splitterItem2.AllowHotTrack = True
            splitterItem2.Location = New System.Drawing.Point(0, 295)
            splitterItem2.Name = "splitterItem2"
            splitterItem2.Size = New System.Drawing.Size(813, 10)
            ' 
            ' layoutControlItem4
            ' 
            layoutControlItem4.Control = displayResultControl1
            layoutControlItem4.Location = New System.Drawing.Point(0, 305)
            layoutControlItem4.Name = "layoutControlItem4"
            layoutControlItem4.Size = New System.Drawing.Size(813, 343)
            layoutControlItem4.TextSize = New System.Drawing.Size(0, 0)
            layoutControlItem4.TextVisible = False
            ' 
            ' layoutControlItem2
            ' 
            layoutControlItem2.ContentVertAlignment = DevExpress.Utils.VertAlignment.Center
            layoutControlItem2.Control = checkEdit1
            layoutControlItem2.Location = New System.Drawing.Point(584, 0)
            layoutControlItem2.Name = "layoutControlItem2"
            layoutControlItem2.Size = New System.Drawing.Size(229, 36)
            layoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
            layoutControlItem2.TextVisible = False
            ' 
            ' layoutControlItem3
            ' 
            layoutControlItem3.Control = xtraTabControl1
            layoutControlItem3.Location = New System.Drawing.Point(0, 36)
            layoutControlItem3.Name = "layoutControlItem3"
            layoutControlItem3.Size = New System.Drawing.Size(813, 259)
            layoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
            layoutControlItem3.TextVisible = False
            ' 
            ' codeExampleNameLbl
            ' 
            codeExampleNameLbl.AllowHotTrack = False
            codeExampleNameLbl.AppearanceItemCaption.Font = New System.Drawing.Font("Arial", 20.25F)
            codeExampleNameLbl.AppearanceItemCaption.Options.UseFont = True
            codeExampleNameLbl.Location = New System.Drawing.Point(0, 0)
            codeExampleNameLbl.MinSize = New System.Drawing.Size(100, 36)
            codeExampleNameLbl.Name = "codeExampleNameLbl"
            codeExampleNameLbl.Size = New System.Drawing.Size(584, 36)
            codeExampleNameLbl.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
            codeExampleNameLbl.TextSize = New System.Drawing.Size(335, 32)
            ' 
            ' Form1
            ' 
            AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New System.Drawing.Size(1248, 668)
            Me.Controls.Add(layoutControl1)
            Name = "Form1"
            CType(checkEdit1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
            layoutControl1.ResumeLayout(False)
            CType(treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            xtraTabControl1.ResumeLayout(False)
            xtraTabPage1.ResumeLayout(False)
            xtraTabPage2.ResumeLayout(False)
            xtraTabPage3.ResumeLayout(False)
            xtraTabPage4.ResumeLayout(False)
            CType(Root, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
            CType(splitterItem1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(splitterItem2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(layoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
            CType(codeExampleNameLbl, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
        End Sub

#End Region
        Private codeEditor As ExampleCodeEditor

        Private evaluator As ExampleEvaluatorByTimer

        Private examples As List(Of CodeExampleGroup)

        Private checkEdit1 As CheckEdit

        Private xtraTabPage3 As XtraTabPage

        Private richEditControlCSClass As RichEditControl

        Private xtraTabPage4 As XtraTabPage

        Private richEditControlVBClass As RichEditControl

        Private treeListRootNodeLoading As Boolean = True

        Private richEditControl As RichEditControl

        Public Sub New()
            InitializeComponent()
            InitializeRichEditControl()
            Dim examplePath As String = GetExamplePath("CodeExamples")
            Dim examplesCS As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp)
            Dim examplesVB As Dictionary(Of String, FileInfo) = GatherExamplesFromProject(examplePath, ExampleLanguage.VB)
            DisableTabs(examplesCS.Count, examplesVB.Count)
            examples = FindExamples(examplePath, examplesCS, examplesVB)
            ShowExamplesInTreeList(treeList1, examples)
            codeEditor = New ExampleCodeEditor(richEditControlCS, richEditControlVB, richEditControlCSClass, richEditControlVBClass)
            CurrentExampleLanguage = DetectExampleLanguage("RichEditAPISample")
            evaluator = New RichEditExampleEvaluatorByTimer()
            AddHandler evaluator.QueryEvaluate, AddressOf OnExampleEvaluatorQueryEvaluate
            AddHandler evaluator.OnBeforeCompile, AddressOf evaluator_OnBeforeCompile
            AddHandler evaluator.OnAfterCompile, AddressOf evaluator_OnAfterCompile
            AddHandler xtraTabControl1.SelectedPageChanged, AddressOf xtraTabControl1_SelectedPageChanged
            ShowFirstExample("Range")
            treeList1.CollapseAll()
        End Sub

        Private Sub InitializeRichEditControl()
            richEditControl = displayResultControl1.RichEdit
        End Sub

        Public Property CurrentExampleLanguage As ExampleLanguage
            Get
                If Equals(xtraTabControl1.SelectedTabPage.Tag.ToString(), "CS") Then
                    Return ExampleLanguage.Csharp
                Else
                    Return ExampleLanguage.VB
                End If
            End Get

            Set(ByVal value As ExampleLanguage)
                codeEditor.CurrentExampleLanguage = value
            'xtraTabControl1.SelectedTabPageIndex = (value == ExampleLanguage.Csharp) ? 0 : 1;
            End Set
        End Property

        Private Sub ShowExamplesInTreeList(ByVal treeList As TreeList, ByVal examples As List(Of CodeExampleGroup))
#Region "InitializeTreeList"
            treeList.OptionsPrint.UsePrintStyles = True
            AddHandler treeList.FocusedNodeChanged, New FocusedNodeChangedEventHandler(AddressOf OnNewExampleSelected)
            treeList.OptionsView.ShowColumns = False
            treeList.OptionsView.ShowIndicator = False
            AddHandler treeList.VirtualTreeGetChildNodes, AddressOf treeList_VirtualTreeGetChildNodes
            AddHandler treeList.VirtualTreeGetCellValue, AddressOf treeList_VirtualTreeGetCellValue
#End Region
            Dim col1 As TreeListColumn = New TreeListColumn()
            col1.VisibleIndex = 0
            col1.OptionsColumn.AllowEdit = False
            col1.OptionsColumn.AllowMove = False
            col1.OptionsColumn.ReadOnly = True
            treeList.Columns.AddRange(New TreeListColumn() {col1})
            treeList.DataSource = New [Object]()
            treeList.ExpandAll()
        End Sub

        Private Sub treeList_VirtualTreeGetCellValue(ByVal sender As Object, ByVal args As VirtualTreeGetCellValueInfo)
            Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
            If group IsNot Nothing Then args.CellData = group.Name
            Dim example As CodeExample = TryCast(args.Node, CodeExample)
            If example IsNot Nothing Then args.CellData = example.RegionName
        End Sub

        Private Sub treeList_VirtualTreeGetChildNodes(ByVal sender As Object, ByVal args As VirtualTreeGetChildNodesInfo)
            If treeListRootNodeLoading Then
                args.Children = examples
                treeListRootNodeLoading = False
            Else
                If args.Node Is Nothing Then Return
                Dim group As CodeExampleGroup = TryCast(args.Node, CodeExampleGroup)
                If group IsNot Nothing Then args.Children = group.Examples
            End If
        End Sub

        Private Sub ShowFirstExample(ByVal firstGroupName As String)
            treeList1.ExpandAll()
            If treeList1.Nodes.Count > 0 Then treeList1.FocusedNode = treeList1.FindNodeByFieldValue("", firstGroupName).NextVisibleNode
        End Sub

        Private Sub evaluator_OnAfterCompile(ByVal sender As Object, ByVal args As OnAfterCompileEventArgs)
            codeEditor.AfterCompile(args.Result)
        End Sub

        Private Sub evaluator_OnBeforeCompile(ByVal sender As Object, ByVal e As EventArgs)
            Dim document As Document = richEditControl.Document
            document.BeginUpdate()
            codeEditor.BeforeCompile()
            richEditControl.CreateNewDocument()
            document.Unit = DevExpress.Office.DocumentUnit.Document
            document.EndUpdate()
        End Sub

        Private Sub OnNewExampleSelected(ByVal sender As Object, ByVal e As FocusedNodeChangedEventArgs)
            Dim newExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.Node), CodeExample)
            Dim oldExample As CodeExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.OldNode), CodeExample)
            If newExample Is Nothing Then Return
            Dim exampleCode As String = codeEditor.ShowExample(oldExample, newExample)
            codeExampleNameLbl.Text = ConvertStringToMoreHumanReadableForm(newExample.RegionName) & " example"
            Dim args As CodeEvaluationEventArgs = New CodeEvaluationEventArgs()
            InitializeCodeEvaluationEventArgs(args)
            evaluator.ForceCompile(args)
            If Equals(newExample.HumanReadableGroupName, "Comments") Then
                richEditControl.Options.Comments.Visibility = RichEditCommentVisibility.Visible
                displayResultControl1.DockPanel.Show()
            Else
                richEditControl.Options.Comments.Visibility = RichEditCommentVisibility.Hidden
                displayResultControl1.DockPanel.Hide()
            End If
        End Sub

        Private Sub InitializeCodeEvaluationEventArgs(ByVal e As CodeEvaluationEventArgs)
            e.Result = True
            e.Code = codeEditor.CurrentCodeEditor.Text
            e.CodeClasses = codeEditor.CurrentCodeClassEditor.Text
            e.Language = CurrentExampleLanguage
            e.EvaluationParameter = richEditControl.Document
        End Sub

        Private Sub OnExampleEvaluatorQueryEvaluate(ByVal sender As Object, ByVal e As CodeEvaluationEventArgs)
            e.Result = False
            If codeEditor.RichEditTextChanged Then ' && compileComplete) {
                Dim span As TimeSpan = Date.Now - codeEditor.LastExampleCodeModifiedTime
                If span < TimeSpan.FromMilliseconds(1000) Then 'CompileTimeIntervalInMilliseconds  1900
                    codeEditor.ResetLastExampleModifiedTime()
                    Return
                End If

                'e.Result = true;
                InitializeCodeEvaluationEventArgs(e)
            End If
        End Sub

        Private Sub DisableTabs(ByVal examplesCSCount As Integer, ByVal examplesVBCount As Integer)
            If examplesCSCount = 0 Then
                For Each t As XtraTabPage In xtraTabControl1.TabPages
                    If Equals(t.Tag.ToString(), "CS") Then t.PageEnabled = False
                Next
            End If

            If examplesVBCount = 0 Then
                For Each t As XtraTabPage In xtraTabControl1.TabPages
                    If Equals(t.Tag.ToString(), "VB") Then t.PageEnabled = False
                Next
            End If
        End Sub

        Private Sub checkEdit1_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            If checkEdit1.Checked Then
                AddHandler richEditControl.MouseMove, AddressOf richEditControl_MouseMove
            Else
                RemoveHandler richEditControl.MouseMove, AddressOf richEditControl_MouseMove
            End If
        End Sub

        Private Sub xtraTabControl1_SelectedPageChanged(ByVal sender As Object, ByVal e As TabPageChangedEventArgs)
            CurrentExampleLanguage = If((Equals(e.Page.Tag.ToString(), "CS")), ExampleLanguage.Csharp, ExampleLanguage.VB)
        End Sub

#Region "#getpositionfrrompoint"
        Private Sub richEditControl_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim docPoint As Point = Units.PixelsToDocuments(e.Location, richEditControl.DpiX, richEditControl.DpiY)
            Dim pos As DocumentPosition = richEditControl.GetPositionFromPoint(docPoint)
            If pos IsNot Nothing Then
                Text = String.Format("Mouse is over position {0}", pos)
            Else
                Text = ""
            End If
        End Sub
#End Region  ' #getpositionfrrompoint
    End Class
End Namespace
