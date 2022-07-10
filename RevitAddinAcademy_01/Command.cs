#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Hardcoding file path temporarily
            // Path slashes - C:\\Path\\ or @"C:\Path\
            string excelFile = @"C:\\Users\\David\\Documents\SheetIndex_00000.xlsx";
            
            // define excel app
            Excel.Application excelApp = new Excel.Application();
            // open workbook (file)
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            // open worksheet (tab) item 1 is first worksheet
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[1];

            // select all cells in use (usedrange)
            Excel.Range excelRng = excelWs.UsedRange;

            int rowCount = excelRng.Rows.Count;

            // do some stuff in excel

            List<string[]> dataList = new List<string[]>();

            // excel imports as 2D array
            for(int i = 1; i <= rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1];
                Excel.Range cell2 = excelWs.Cells[i, 2];
                Excel.Range cell3 = excelWs.Cells[i, 3];
                Excel.Range cell4 = excelWs.Cells[i, 4];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();
                string data3 = cell3.Value.ToString();
                string data4 = cell4.Value.ToString();


                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                dataList.Add(dataArray);

            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Stuff");
                
                Level curLevel = Level.Create(doc, 100);


                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                curSheet.SheetNumber = "A1";
                curSheet.Name = "New Sheet";


                t.Commit();

            }

            


            excelWb.Close();
            excelApp.Quit();



            return Result.Succeeded;
        }
    }
}
