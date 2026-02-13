<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128612654/16.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T480591)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/WpfSpreadsheet_BindToDataSource/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/WpfSpreadsheet_BindToDataSource/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/WpfSpreadsheet_BindToDataSource/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/WpfSpreadsheet_BindToDataSource/MainWindow.xaml.vb))
<!-- default file list end -->
# How to bind a spreadsheet to an MS SQL Server database (WPF Spreadsheet)


This example demonstrates how to bind a cell range on a worksheet to the sample <strong>Northwind</strong> database to load data from the <strong>Suppliers</strong> data table. To accomplish this task, the <a href="https://documentation.devexpress.com/#CoreLibraries/DevExpressSpreadsheetWorksheetDataBindingCollection_BindToDataSourcetopic">WorksheetDataBindingCollection.BindToDataSource</a> method is used.<br>This application also enables end-users to add, modify or remove data in a data table. They can use the corresponding buttons on the <strong>File</strong> tab, in the <strong>Database</strong> group to edit the data and save their changes back to the database. <br>To insert new rows, a data entry form is used. The user should fill out the given data entry fields and click the <strong>Save </strong>cell to add a new record to the <strong>Suppliers </strong>data table. Clicking the <strong>Apply Changes </strong>button posts the updated data back to the database. To remove a record, the user should select the required Suppliers row on the worksheet and click the <strong>Remove Record </strong>button. The <strong>Delete</strong> dialog will be invoked asking the user to confirm the delete operation. <br>To send the modified data to the connected database, the <strong>Update</strong> method of the <strong>SuppliersTableAdapter</strong> is used. <br><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-bind-a-spreadsheet-to-an-ms-sql-server-database-wpf-spreadsheet-t480591/16.2.3+/media/03d39ba1-edde-11e6-80bf-00155d62480c.png">

<br/>


<!-- feedback -->
## Does This Example Address Your Development Requirements/Objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-spreadsheet-bind-spreadsheet-to-ms-sql-server-database&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=wpf-spreadsheet-bind-spreadsheet-to-ms-sql-server-database&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
