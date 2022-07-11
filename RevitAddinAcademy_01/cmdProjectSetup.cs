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
    public class cmdProjectSetup : IExternalCommand
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

            //Hardcoded Path
            string excelFile = @"C:\\Users\\David\\Documents\\Addin Academy\\Session02_Challenge-220706-113155.xlsx";

            //Define excel app and open workbook, worksheets 1 AND 2
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWs2 = excelWb.Worksheets.Item[2];
            Excel.Range excelRng1 = excelWs1.UsedRange;
            Excel.Range excelRng2 = excelWs2.UsedRange;

            //Select all cells in use            
            List<string[]> levelList = new List<string[]>();
            List<string[]> sheetList = new List<string[]>();
            int rowCount1 = excelRng1.Rows.Count;
            int rowCount2 = excelRng2.Rows.Count;

            // Get Sheet Type (first one)
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector.WhereElementIsElementType();
            var tbType = collector.FirstElementId();

            // Should make this a method and instance the columns maybe
            // Or just derive sheet names and numbers from Level Data
            //Import excel LEVEL data as 2D array
            for (int i = 1; i <= rowCount1; i++)
            {
                Excel.Range cell1 = excelWs1.Cells[i, 1];
                Excel.Range cell2 = excelWs1.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                levelList.Add(dataArray);                

            }

            for (int i = 1; i <= rowCount2; i++)
            {
                Excel.Range cell1 = excelWs2.Cells[i, 1];
                Excel.Range cell2 = excelWs2.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                sheetList.Add(dataArray);

            }

            //Remove Headers
            levelList.RemoveAt(0);
            sheetList.RemoveAt(0);

            //Create Level from data
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels");

                //Level curLevel = Level.Create(doc, 100);
                foreach (string[] levelData in levelList)
                {
                    string levelName = levelData[0];
                    string levelFeetstr = levelData[1];
                    double levelFeet = Double.Parse(levelFeetstr);
                    Level curLevel = Level.Create(doc, levelFeet);
                    // Need to Rename existing Levels too                    
                    string tempName = "Element " + curLevel.Id.ToString();
                    curLevel.Name = tempName;
                    curLevel.Name = levelName;

                }

                t.Commit();

            }

            //Create Sheet from data            
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Sheets");

                foreach (string[] sheetData in sheetList)
                {
                    string sheetNum = sheetData[0];
                    string sheetName = sheetData[1];
                    ViewSheet curSheet = ViewSheet.Create(doc, tbType);
                    curSheet.SheetNumber = sheetNum;
                    curSheet.Name = sheetName;

                    // Need to check for existing sheets


                }

                t.Commit();

            }
            
            excelWb.Close();
            excelApp.Quit();


            return Result.Succeeded;
        }
    }
}
