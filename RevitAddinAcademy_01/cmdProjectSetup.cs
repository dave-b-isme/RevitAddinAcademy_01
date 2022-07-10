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
            List<string[]> dataList = new List<string[]>();
            int rowCount = excelRng1.Rows.Count;

            //Import excel LEVEL data as 2D array
            for(int i = 1; i <= rowCount; i++)
            {
                Excel.Range cell1 = excelWs1.Cells[i, 1];
                Excel.Range cell2 = excelWs1.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                dataList.Add(dataArray);                

            }

            //Remove Headers
            dataList.RemoveAt(0);

            //Create Level and Sheet from data
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Levels");

                //Level curLevel = Level.Create(doc, 100);
                foreach (string[] levelHeight in dataList)
                {
                    string levelName = levelHeight[0];
                    string levelFeetstr = levelHeight[1];
                    double levelFeet = Double.Parse(levelFeetstr);
                    Level curLevel = Level.Create(doc, levelFeet);

                }

                t.Commit();

            }
            
            excelWb.Close();
            excelApp.Quit();


            return Result.Succeeded;
        }
    }
}
