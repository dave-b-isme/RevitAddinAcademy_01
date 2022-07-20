#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;
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

            TaskDialog.Show("Revit Addin Academy", "Hello There");

            // Select the Excel file with FileBrowserDia
            //string filePath = @"C:\Users\David\Documents\Addin Academy\Session03_Challenge-220713-114007.xlsx";
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\Users\David\Documents";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls; *.xlsx; *.xlsm | All Files | *.*";            
            
            //GATE Break code with user input
            if(dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return Result.Failed;                
            }
            
            string filePath = dialog.FileName;

            //
            //REVIEW Use Struct with Method to get Level & Sheet Data
            //
            List<string[]> levelArrays = GetExcelData(filePath, "Levels");
            List<LevelStruct> levelData = GetLevelData(levelArrays);
            List<string[]> sheetArray = GetExcelData(filePath, "Sheets");
            List<SheetStruct> sheetData = GetSheetData(sheetArray);

            // Getting View Types...
            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            collector1.OfClass(typeof(ViewFamilyType));
            ViewFamilyType curVFT = null;
            ViewFamilyType curRCPVFT = null;

            foreach (ViewFamilyType curElem in collector1)
            {
                if (curElem.ViewFamily == ViewFamily.FloorPlan)
                {
                    curVFT = curElem;
                }

                else
                    if (curElem.ViewFamily == ViewFamily.CeilingPlan)
                {
                    curRCPVFT = curElem;
                }

            }


            // REVIEW Challenge attempt
            // Create Levels, Floor Plan and RCP Views for Each Level
            // TRANSACTION

            int levelCounter = 0;
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Making Plans");
                
                foreach (LevelStruct level in levelData)
                {
                    Level curLevel = null;
                    curLevel = Level.Create(doc, (level.levelElev));
                    curLevel.Name = level.levelName + curLevel.Id.ToString();

                    try
                    {
                        curLevel.Name = level.levelName;
                        levelCounter++;
                    }
                    catch (Exception ex)
                    {
                        Debug.Print(ex.Message);
                        TaskDialog.Show("Revit Addin Academy", "Level" + level.levelName + "Has been renamed to " + level.levelName + curLevel.Id.ToString());
                        //levelData.Remove(level);
                    }

                    ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, curLevel.Id);
                    ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, curLevel.Id);
                    curRCP.Name = curRCP.Name + " RCP";

                }

                t.Commit();
            }

            // REVIEW Challenge attempt
            // Create Sheets
            string tbString = "E1 30x42 Horizontal";

            using (Transaction t2 = new Transaction(doc))
            {
                t2.Start("Making Sheets");

                Element tbType = GetTBbyName(doc, tbString);

                if(tbType == null)
                {
                    TaskDialog.Show("Error", "Oops can't find Title Block");
                    return Result.Failed;
                }

                foreach(SheetStruct sheet in sheetData)
                {
                    ViewSheet newSheet = null;
                    
                    try
                    {
                        newSheet = ViewSheet.Create(doc, tbType.Id);
                        newSheet.Name = sheet.sheetName;
                        newSheet.SheetNumber = sheet.sheetNumber;
                        SetParameter(newSheet, "Drawn By", sheet.drawnBy);
                        SetParameter(newSheet, "Checked By", sheet.checkBy);
                    }
                    catch (Exception ex)
                    {
                        Debug.Print(ex.Message);
                    }

                    try
                    {
                        Viewport newViewport = Viewport.Create(doc, newSheet.Id, GetViewByName(doc, sheet.sheetView), new XYZ(0, 0, 0));
                    }
                    catch(Exception ex2)
                    {
                        Debug.Print(ex2.Message);
                    }
                    
                    


                }




                t2.Commit();
            }

            //// First Challenge attempt
            //// Get the excel data (by name) with method - GetExcelData ( filepath, sheetnames ) Levels Sheets            
            //List<string[]> lvlData = GetExcelData(filePath, "Levels");
            //int levelCount = lvlData.Count;
            //var levelArray = lvlData.ToArray();
            //List<string[]> shtData = GetExcelData(filePath, "Sheets");
            //int sheetCount = shtData.Count;
            //var sheetArray = shtData.ToArray();
            //int levelCounter = 0;


            return Result.Succeeded;
        }

        private ElementId GetViewByName(Document doc, string sheetView)
        {
            FilteredElementCollector collector3 = new FilteredElementCollector(doc);
            collector3.OfCategory(BuiltInCategory.OST_Views);

            foreach(View v in collector3)
            {
                if(v.Name == sheetView)
                {
                    return v.Id;
                }

            }

            return null;
        }

        //STRUCTS
        //
        // LEVEL DATA
        private struct LevelStruct
        {
            public string levelName;
            public double levelElev;

            public LevelStruct(string name, double elev)
            {
                levelName = name;
                levelElev = elev;
            }

        }

        //
        // SHEET DATA

        private struct SheetStruct
        {
            public string sheetNumber;
            public string sheetName;
            public string sheetView;
            public string drawnBy;
            public string checkBy;

            public SheetStruct(string number, string name, string view, string db, string cb)
            {
                sheetNumber = number;
                sheetName = name;
                sheetView = view;
                drawnBy = db;
                checkBy = cb;
            }
        }


        // METHODS
        //
        // GET TITLE BLOCK BY NAME
        internal Element GetTBbyName(Document doc, string tbName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType();
            
            foreach (Element e in collector)
            {
                if (e.Name == tbName)
                {
                    return e;
                }
                else
                {
                    Element wrongE = collector.FirstElement();
                    return wrongE;
                }
                
            }

            return null;

        }

        //
        // SET PARAMETER VALUE BY NAME
        // doc, element, ParameterName, Parameter value
        internal Parameter SetParameter(Element e, string pName, string pValue)
        {
            foreach (Parameter curParam in e.Parameters)
            {
                if(curParam.Definition.Name == pName)
                {
                    curParam.Set(pValue);
                }
            }

            return null;
        }

        // GET EXCEL WORKSHEET BY NAME
        // AND
        // GET EXCEL DATA FROM WORKSHEET
        // AND
        // CLOSE EXCEL
        //
        // Worksheet Name, string array
        // Try 2D array instead of List?
        // Try Struct instead of List?

        
        // METHOD
        // Get worksheet by name
        // Trying one method to get worksheet of any size to use for both sheets and levels
        // ADD use "Sheet1" if Name Not Found
        // ADD try catch if no sheets found
        internal List<string[]> GetExcelData(string excelFile, string wsName)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWS = excelWB.Worksheets[wsName];
            Excel.Range excelRng = excelWS.UsedRange;

            int colCount = excelRng.Columns.Count;
            int rowCount = excelRng.Rows.Count;

            
            List<string[]> dataList = new List<string[]>();

            for (int r = 2; r <= rowCount; r++)
            {
                string[] dataArray = new string[colCount];

                for (int c = 1; c <= colCount; c++)
                {
                    Excel.Range cell = excelWS.Cells[r, c];
                    dataArray[c-1] = cell.Value.ToString();

                }

                dataList.Add(dataArray);

            }

            excelWB.Close();
            excelApp.Quit();

            return dataList;
        }

        // Do I need to loop through or is there a way to convert List<string[]> to LevelStruct
        // Maybe create LevelStruct as a List<string[]>??
        // Maybe collect data into LevelStruct XXXX no, would need to define # of columns in Struct
        private List<LevelStruct> GetLevelData(List<string[]> curData)
        {
            List<LevelStruct> returnList = new List<LevelStruct>();

            //foreach(LevelStruct Data in curData)
            foreach (string[] data in curData)
            {
                //LevelStruct curLevel = new LevelStruct(data);
                //It won't let me plug array into Struct so I index it 
                LevelStruct curLevel = new LevelStruct(data[0], Double.Parse(data[1]));
                returnList.Add(curLevel);
            }

            return returnList;
        }
                
        private List<SheetStruct> GetSheetData(List<string[]> curData)
        {
            List<SheetStruct> returnList = new List<SheetStruct>();

            foreach (string[] data in curData)
            {
                SheetStruct curSheet = new SheetStruct(data[0], data[1], data[2], data[3], data[4]);
                returnList.Add(curSheet);
            }

            return returnList;
        }

    }
}
