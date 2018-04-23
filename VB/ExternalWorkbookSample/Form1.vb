Imports DevExpress.Spreadsheet
Imports System
Imports System.Data
Imports System.Windows.Forms

Namespace ExternalWorkbookSample
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnAddExternalWorkbook_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddExternalWorkbook.ItemClick
'            #Region "#addexternalworkbook"
            Dim externalWorkbook As New Workbook()
            externalWorkbook.Worksheets(0).Import(CreateDataTable(10), False, 0, 0)
            externalWorkbook.Options.Save.CurrentFileName = "ExternalDocument.xlsx"
            For Each item As IWorkbook In spreadsheetControl1.Document.ExternalWorkbooks
                If item.Options.Save.CurrentFileName = externalWorkbook.Options.Save.CurrentFileName Then
                    Return
                End If
            Next item
            spreadsheetControl1.Document.ExternalWorkbooks.Add(externalWorkbook)
'            #End Region ' #addexternalworkbook
            MessageBox.Show(String.Format("The external workbook {0} is added. Now you can reference it.", externalWorkbook.Options.Save.CurrentFileName))
        End Sub

        Private Sub btnInsertExternalReference_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnInsertExternalReference.ItemClick
'            #Region "#insertexternalreference"
            If spreadsheetControl1.Document.ExternalWorkbooks.Count = 0 Then
                MessageBox.Show("Add a workbook to the ExternalWorkbooks collection")
                Return
            End If
            Dim extWorkbook As IWorkbook = DirectCast(spreadsheetControl1.Document.ExternalWorkbooks(0), IWorkbook)
            Dim extWorkbookName As String = extWorkbook.Options.Save.CurrentFileName
            Dim sFormula As String = String.Format("=[{0}]Sheet1!A1", extWorkbookName)
            spreadsheetControl1.Document.Worksheets(0).Cells("A1").Formula = sFormula
'            #End Region ' #insertexternalreference
        End Sub

        Private Function CreateDataTable(ByVal rowCount As Integer) As DataTable
            Dim someDT As New DataTable()
            For i As Integer = 0 To 4
                someDT.Columns.Add("Value" & i.ToString(), GetType(Integer))
            Next i
            Dim myRand As New Random()
            For i As Integer = 0 To rowCount - 1
                someDT.Rows.Add(myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100))
            Next i
            Return someDT

        End Function
    End Class
End Namespace
