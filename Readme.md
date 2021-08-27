<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613888/14.1.6%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T144947)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/ExternalWorkbookSample/Form1.cs) (VB: [Form1.vb](./VB/ExternalWorkbookSample/Form1.vb))
<!-- default file list end -->
# How to reference an external workbook


This example demonstrates how to insert an external reference link from a workbook to another workbook. <br />An external workbook is created at runtime. Note that use of this method in production code requires a license to the DevExpress Document Server or the DevExpress Universal Subscription and a reference to the DevExpress.Docs.vX.Y.dll assembly. You can use a workbook loaded into another SpreadsheetControl instead (the <strong>SpreadsheetControl.Document</strong> property).<br />The external workbook is populated with random data by importing a data table at runtime.<br />Subsequently the workbook is added to the <a href="http://help.devexpress.com/#CoreLibraries/clsDevExpressSpreadsheetExternalWorkbookCollectiontopic">ExternalWorkbookCollection</a>. <br />A cell formula in the current worksheet can reference a workbook contained in that collection.<br /><br />See also the <a href="http://help.devexpress.com/#WindowsForms/CustomDocument12272">Cell Referencing</a> documentation topic available online.<br /><br />

<br/>


