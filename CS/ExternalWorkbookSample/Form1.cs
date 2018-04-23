using DevExpress.Spreadsheet;
using System;
using System.Data;
using System.Windows.Forms;

namespace ExternalWorkbookSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnAddExternalWorkbook_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #addexternalworkbook
            Workbook externalWorkbook = new Workbook();
            externalWorkbook.Worksheets[0].Import(CreateDataTable(10), false, 0, 0);
            externalWorkbook.Options.Save.CurrentFileName = "ExternalDocument.xlsx";
            foreach (IWorkbook item in spreadsheetControl1.Document.ExternalWorkbooks)
            {
                if (item.Options.Save.CurrentFileName == externalWorkbook.Options.Save.CurrentFileName)
                    return;
            }
            spreadsheetControl1.Document.ExternalWorkbooks.Add(externalWorkbook);
            #endregion #addexternalworkbook
            MessageBox.Show(String.Format("The external workbook {0} is added. Now you can reference it.", externalWorkbook.Options.Save.CurrentFileName));
        }

        private void btnInsertExternalReference_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #insertexternalreference
            if (spreadsheetControl1.Document.ExternalWorkbooks.Count == 0)
            {
                MessageBox.Show("Add a workbook to the ExternalWorkbooks collection");
                return;
            }
            IWorkbook extWorkbook = (IWorkbook) spreadsheetControl1.Document.ExternalWorkbooks[0];
            string extWorkbookName = extWorkbook.Options.Save.CurrentFileName;
            string sFormula = String.Format("=[{0}]Sheet1!A1", extWorkbookName);
            spreadsheetControl1.Document.Worksheets[0].Cells["A1"].Formula = sFormula;
            #endregion #insertexternalreference
        }

        DataTable CreateDataTable(int rowCount)
        {
            DataTable someDT = new DataTable();
            for (int i = 0; i < 5; i++)
            {
                someDT.Columns.Add("Value" + i.ToString(), typeof(int));
            }
            Random myRand = new Random();
            for (int i = 0; i < rowCount; i++)
            {
                someDT.Rows.Add(myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100), myRand.Next(1, 100));
            }
            return someDT;

        }
    }
}
