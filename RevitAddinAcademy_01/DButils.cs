using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
//using Forms = System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
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

    public class Furniture
    {
        public List<FurnitureSet> FurnitureList { get; set; }

        public Furniture(List<FurnitureSet> furnList)
        {
            FurnitureList = furnList;
        }
        
        public int GetFurnitureCount()
        {
            return FurnitureList.Count;
        }

    }

    // UTILITIES

    public static class Util
    {
        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            FilteredElementCollector coll = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms);
            List<SpatialElement> roomList = new List<SpatialElement>();

            foreach(Element curElem in coll)
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

            foreach(Element e in coll)
            {
                Family curFam = e as Family;

                if(curFam.Name == famName)
                {
                    ISet<ElementId> famTypeIdList = curFam.GetFamilySymbolIds();
                    foreach(ElementId famTypeId in famTypeIdList)
                    {
                        FamilySymbol curFS = doc.GetElement(famTypeId) as FamilySymbol;
                        if(curFS.Name == typeName)
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
                if(curParam.Definition.Name == paramName)
                    return curParam.AsString();
            }
            return null;
        }

        public static void SetParameter(Element curElem, string paramName, string paramValue)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if(curParam.Definition.Name == paramName)
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
