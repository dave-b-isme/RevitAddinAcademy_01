﻿using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Architecture;
//using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.DB.Electrical;
//using Autodesk.Revit.DB.Mechanical;
//using Autodesk.Revit.DB.Plumbing;

namespace RevitAddinAcademy_01
{
    internal class DButils
    {
    }

    public class FurnitureSet
    {
        public string SetName { get; set; }
        public string RoomName { get; set; }
        public List<string> FurnSet { get; set; }

        public FurnitureSet(string setName, string roomName, string furnSet)
        {
            SetName = setName;
            RoomName = roomName;
            FurnSet = FormatSetList(furnSet);
        }

        private List<string> FormatSetList(string furnName)
        {
            List<string> returnList = furnName.Split(',').ToList();
            return returnList;
        }
    }

    public class FurnitureList
    {
        public List<FurnitureSet> FurnList { get; set; }

        public FurnitureList(List<FurnitureSet> furnList)
        {
            FurnList = furnList;
        }
    }

    public class FurnitureType
    {
        public string Name { get; set; }
        public string FamName { get; set; }
        public string FamType { get; set; }

        public FurnitureType(string name, string famName, string famType)
        {
            Name = name;
            FamName = famName;
            FamType = famType;
        }

    }
    public class FurnitureTypes
    {
        public List<FurnitureType> FurnTypes { get; set; }
        public FurnitureTypes(List<FurnitureType> furnTypes)
        {
            FurnTypes = furnTypes;
        }
        public int GetFurnitureCount()
        {
            return FurnTypes.Count;
        }

    }

    // UTILITIES

    // Try a GetTypeByName() with input string "WallType" or "LineType" etc
    // switch statement with case string = walltype to get WallType as element etc.
    // WallType is ElementType so maybe send that instead of string?

    // return (workset != null) ? workset : worksets.FirstOrDefault();

    public static class Util
    {
        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            FilteredElementCollector coll = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms);
            List<SpatialElement> roomList = new List<SpatialElement>();

            foreach (Element curElem in coll)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                roomList.Add(curRoom);
            }
            return roomList;
        }

        public static FamilySymbol GetFamTypeByName(Document doc, string famName, string typeName)
        {
            FilteredElementCollector coll = new FilteredElementCollector(doc)
                .OfClass(typeof(Family));

            foreach (Element e in coll)
            {
                Family curFam = e as Family;

                if (curFam.Name == famName)
                {
                    ISet<ElementId> famTypeIdList = curFam.GetFamilySymbolIds();
                    foreach (ElementId famTypeId in famTypeIdList)
                    {
                        FamilySymbol curFS = doc.GetElement(famTypeId) as FamilySymbol;
                        if (curFS.Name == typeName)
                            return curFS;
                    }
                }
            }
            return null;
        }

        public static string GetParameter(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    return curParam.AsString();
            }
            return null;
        }

        public static void SetParameter(Element curElem, string paramName, string paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }
        public static void SetParameter(Element curElem, string paramName, double paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }
        public static void SetParameter(Element curElem, string paramName, int paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }

        // Maybe I don't really want to break these up

        public static List<string[]> GetExcelWS(string wsName)
        {
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            //dialog.InitialDirectory = @"C:\Users\David\Documents";
            dialog.InitialDirectory = @"%USERPROFILE%\Documents";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls; *.xlsx; *.xlsm | All Files | *.*";

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return null;
            }

            string filePath = dialog.FileName;

            return GetExcelWSdata(filePath, wsName);

        }

        public static List<string[]> GetExcelWS(int wsIndex)
        {
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            //dialog.InitialDirectory = @"C:\Users\David\Documents";
            dialog.InitialDirectory = @"%USERPROFILE%\Documents";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls; *.xlsx; *.xlsm | All Files | *.*";

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return null;
            }

            string filePath = dialog.FileName;

            return GetExcelWSdata(filePath, wsIndex);

        }

        public static List<string[]> GetExcelWSdata(string filePath, string wsName)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet excelWS = excelWB.Worksheets[wsName];
            Excel.Range excelRng = excelWS.UsedRange;

            int colCount = excelRng.Columns.Count;
            int rowCount = excelRng.Rows.Count;


            List<string[]> dataList = new List<string[]>();

            for (int r = 1; r <= rowCount; r++)
            {
                string[] dataArray = new string[colCount];

                for (int c = 1; c <= colCount; c++)
                {
                    Excel.Range cell = excelWS.Cells[r, c];
                    dataArray[c - 1] = cell.Value.ToString();

                }

                dataList.Add(dataArray);

            }

            excelWB.Close();
            excelApp.Quit();

            return dataList;
        }
        public static List<string[]> GetExcelWSdata(string filePath, int wsIndex)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet excelWS = excelWB.Worksheets[wsIndex];
            Excel.Range excelRng = excelWS.UsedRange;

            int colCount = excelRng.Columns.Count;
            int rowCount = excelRng.Rows.Count;


            List<string[]> dataList = new List<string[]>();

            for (int r = 1; r <= rowCount; r++)
            {
                string[] dataArray = new string[colCount];

                for (int c = 1; c <= colCount; c++)
                {
                    Excel.Range cell = excelWS.Cells[r, c];
                    dataArray[c - 1] = cell.Value.ToString();

                }

                dataList.Add(dataArray);

            }

            excelWB.Close();
            excelApp.Quit();

            return dataList;
        }

        public static List<string[]> GetExcelWBdata(string wsName)
        {
            // Can I get a whole Workbook?
            // Maybe create a class ExcelWB that will take List<string[]>s?
            throw new NotImplementedException();
        }

        public static List<FamilyInstance> GetFurnInRoom(Document doc, Room curRoom)
        {
            FilteredElementCollector furn = new FilteredElementCollector(doc)
                //.OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Furniture)
                .WhereElementIsNotElementType();

            List<FamilyInstance> furnList = new List<FamilyInstance>();

            foreach (Element e in furn)
            {
                FamilyInstance fi = e as FamilyInstance;
                
                if(fi.Room.Id == curRoom.Id)
                {
                    furnList.Add(fi);
                }

            }

            return furnList;
        }


    }


}