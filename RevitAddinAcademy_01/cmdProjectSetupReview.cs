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
    public class cmdProjectSetupReview : IExternalCommand
    {
        // This is the stuff Revit Cares about - methods and structs should sit outside of this
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
            
            //
            //GATE Break code with user input
            //
            if(dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return Result.Failed;                
            }
            
            string filePath = dialog.FileName;

            


            // Get the excel data (by name) with method - GetExcelData ( filepath, sheetnames ) Levels Sheets            
            List<string[]> lvlData = GetExcelData(filePath, "Levels");
            int levelCount = lvlData.Count;
            var levelArray = lvlData.ToArray();
            List<string[]> shtData = GetExcelData(filePath, "Sheets");
            int sheetCount = shtData.Count;
            var sheetArray = shtData.ToArray();
            int levelCounter = 0;
            int sheetCounter = 0;

            using (Transaction t1 = new Transaction(doc))
            {
                foreach( string[] level in lvlData)
                {
                    try
                    {
                        Level curLevel = Level.Create(doc, Double.Parse(level[1]));
                        sheetCounter++;

                    }
                    catch (Exception ex)
                    {
                        Debug.Print(ex.Message);
                    }

                }


                t1.Commit();
            }

            // Create Levels, Floor Plan and RCP Views for Each Level
            // TRANSACTION
            using (Transaction t = new Transaction(doc))
            {
                foreach (string[] levelInst in lvlData)
                {
                    try
                    {
                        Level curLevel = Level.Create(doc, Double.Parse(levelInst[1]));
                        curLevel.Name = "Element " + curLevel.Id.ToString();
                        curLevel.Name = levelInst[0];
                        levelCounter++;
                    }
                    catch (Exception ex)
                    {
                        Debug.Print(ex.Message);
                        TaskDialog.Show("Revit Addin Academy", "Couldn't Create Level" + levelInst[0]);
                    }


                }

                t.Commit();
            }
            



            //Create Sheets and add Views to each sheet
            // TRANSACTION


            
            //string tbName != "E1 30 x 42 Horizontal"; ??? 
            string tbString = "E1 30x42 Horizontal";

            Element tbElem = GetTBbyName(doc, tbString);

            if (tbElem == null)
            {
                TaskDialog.Show("Revit Addin Academy", "I couldn't find any");
            }
            if (tbElem.Name == tbString)
            {
                TaskDialog.Show("Revit Addin Academy", "I found " + tbElem.Name);
            }            
            else
            {
                TaskDialog.Show("Revit Addin Academy", "Can't find it, I'm using " + tbElem.Name + " instead");
            }




            return Result.Succeeded;
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

        // SHEET DATA


        // METHODS
        //
        // GET TITLE BLOCK BY NAME
        private Element GetTBbyName(Document doc, string tbName)
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
        private Parameter SetParameter(Element e, string pName, string pValue)
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

        //Just get worksheet by name
        // Is this the same as ws = curWB.Worksheets[wsName] without loop?

        private Excel.Worksheet GetExcelByName(Excel.Workbook curWb, string wsName)
        {
            foreach (Excel.Worksheet ws in curWb.Worksheets)
            {
                if (ws.Name == wsName)
                {
                    return ws;
                }
            }

            return null;

        }

        private List<LevelStruct> GetLevelData(string excelFile, string wsName)
        {
            List<LevelStruct> dataList = new List<LevelStruct>();
            
            Excel.Application curExcel = new Excel.Application();
            Excel.Workbook curWB = curExcel.Workbooks.Open(excelFile);
            Excel.Worksheet excelWS = curWB.Worksheets[wsName];
            Excel.Range curRng = excelWS.UsedRange;

            int colCount = curRng.Columns.Count;
            int rowCount = curRng.Rows.Count;

            string[] dataArray = new string[colCount];

            for (int r = 2; r<=rowCount; r++)
            {
                for (int c = 1; c<= colCount; c++)
                {
                    Excel.Range cell = excelWS.Cells[r+1, c+1];
                    dataArray[c - 1] = cell.Value.ToString();
                }

                //LevelStruct curLevel = new LevelStruct(dataList.levelName, dataList.levelElev);
                //curLevel.Add(dataArray);

            }


            //return dataList;
            return null;
        }


        private List<string[]> GetExcelData(string excelFile, string wsName)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWS = excelWB.Worksheets[wsName];
            Excel.Range excelRng = excelWS.UsedRange;

            int colCount = excelRng.Columns.Count;
            int rowCount = excelRng.Rows.Count;

            List<string[]> dataList = new List<string[]>();
            string[] dataArray = new string[colCount];

            for (int r = 2; r <= rowCount; r++)
            {
                //string[] colArray = new string[rowCount];
                                
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


    }
}
