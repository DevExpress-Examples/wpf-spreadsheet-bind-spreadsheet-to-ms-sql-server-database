#region #Namespaces
using DevExpress.Spreadsheet;
using System;
using System.Windows;
using WpfSpreadsheet_BindToDataSource.NWindDataSetTableAdapters;
// ...
#endregion #Namespaces

namespace WpfSpreadsheet_BindToDataSource
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DevExpress.Xpf.Ribbon.DXRibbonWindow
    {
        bool applyChangesOnRowsRemoved = false;
        #region #BindToData
        NWindDataSet dataSet;
        SuppliersTableAdapter adapter;
        public MainWindow()
        {
            InitializeComponent();
            BindWorksheetToDataSource();
        }

        private void BindWorksheetToDataSource()
        {
            dataSet = new NWindDataSet();
            adapter = new SuppliersTableAdapter();
            // Populate the "Suppliers" data table with data.
            adapter.Fill(dataSet.Suppliers);

            IWorkbook workbook = spreadsheetControl.Document;
            // Load the template document into the SpreadsheetControl.
            workbook.LoadDocument("Documents\\Suppliers_template.xlsx", DocumentFormat.Xlsx);
            Worksheet sheet = workbook.Worksheets[0];
            // Load data from the "Suppliers" data table into the worksheet starting from the cell "B12".
            sheet.DataBindings.BindToDataSource(dataSet.Suppliers, 11, 1);
        }
        #endregion #BindToData

        #region #UpdateData
        void spreadsheetControl_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Point winPoint = e.GetPosition(spreadsheetControl);
            System.Drawing.Point point = new System.Drawing.Point((int)winPoint.X, (int)winPoint.Y);
            Cell cell = spreadsheetControl.GetCellFromPoint(point);
            if (cell == null)
                return;
            Worksheet sheet = spreadsheetControl.ActiveWorksheet;
            string cellReference = cell.GetReferenceA1();
            // If the "Save" cell is clicked in the data entry form, 
            // add a row containing the entered values to the database table.
            if (cellReference == "I4")
            {
                AddRow(sheet);
                HideDataEntryForm(sheet);
                ApplyChanges();
            }
            // If the "Cancel" cell is clicked in the data entry form, 
            // cancel adding new data and hide the data entry form.
            else if (cellReference == "I6")
            {
                HideDataEntryForm(sheet);
            }
        }

        void AddRow(Worksheet sheet)
        {
            try
            {
                // Append a new row to the "Suppliers" data table.
                dataSet.Suppliers.AddSuppliersRow(
                    sheet["C4"].Value.TextValue, sheet["C6"].Value.TextValue, sheet["C8"].Value.TextValue,
                    sheet["E4"].Value.TextValue, sheet["E6"].Value.TextValue, sheet["E8"].Value.TextValue,
                    sheet.Cells["G4"].DisplayText, sheet.Cells["G6"].DisplayText);
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot add a row to a database table.\n{0}", ex.Message);
                MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void HideDataEntryForm(Worksheet sheet)
        {
            Range range = sheet.Range.Parse("C4,C6,C8,E4,E6,E8,G4,G6");
            range.ClearContents();
            sheet.Rows.Hide(2, 9);
        }

        void ApplyChanges()
        {
            try
            {
                // Send the updated data back to the database.
                adapter.Update(dataSet.Suppliers);
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot update data in a database table.\n{0}", ex.Message);
                MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void spreadsheetControl_RowsRemoving(object sender, RowsChangingEventArgs e)
        {
            Worksheet sheet = spreadsheetControl.ActiveWorksheet;
            Range rowRange = sheet.Range.FromLTRB(0, e.StartIndex, 16383, e.StartIndex + e.Count - 1);
            Range boundRange = sheet.DataBindings[0].Range;
            // If the rows to be removed belong to the data-bound range,
            // display a dialog requesting the user to confirm the deletion of records. 
            if (boundRange.IsIntersecting(rowRange))
            {
                MessageBoxResult result = MessageBox.Show("Want to delete the selected supplier(s)?", "Delete",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                applyChangesOnRowsRemoved = result == MessageBoxResult.Yes;
                e.Cancel = result == MessageBoxResult.No;
                return;
            }
        }
        void spreadsheetControl_RowsRemoved(object sender, RowsChangedEventArgs e)
        {
            if (applyChangesOnRowsRemoved)
            {
                applyChangesOnRowsRemoved = false;
                // Update data in the database.
                ApplyChanges();
            }
        }
        void buttonAddRecord_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            CloseInplaceEditor();
            Worksheet sheet = spreadsheetControl.ActiveWorksheet;
            // Display the data entry form on the worksheet to add a new record to the "Suppliers" data table.
            if (!sheet.Rows[4].Visible)
                sheet.Rows.Unhide(2, 9);
            spreadsheetControl.SelectedCell = sheet["C4"];
        }

        void buttonRemoveRecord_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            CloseInplaceEditor();
            Worksheet sheet = spreadsheetControl.ActiveWorksheet;
            Range selectedRange = spreadsheetControl.Selection;
            Range boundRange = sheet.DataBindings[0].Range;
            // Verify that the selected cell range belongs to the data-bound range.
            if (!boundRange.IsIntersecting(selectedRange) || selectedRange.TopRowIndex < boundRange.TopRowIndex)
            {
                MessageBox.Show("Select a record first!", "Remove Record", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            // Remove the topmost row of the selected cell range.
            sheet.Rows.Remove(selectedRange.TopRowIndex);
        }

        void buttonApplyChanges_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            CloseInplaceEditor();
            // Update data in the database.
            ApplyChanges();
        }

        void buttonCancelChanges_ItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Close the cell in-place editor if it's currently active. 
            CloseInplaceEditor();
            // Load the latest saved data into the "Suppliers" data table.
            adapter.Fill(dataSet.Suppliers);
        }

        void CloseInplaceEditor()
        {
            if (spreadsheetControl.IsCellEditorActive)
                spreadsheetControl.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
        }
        #endregion #UpdateData
    }
}
