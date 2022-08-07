#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
//using Forms = System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.DB.Electrical;
//using Autodesk.Revit.DB.Mechanical;
//using Autodesk.Revit.DB.Plumbing;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class cmdMoveIn : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            //Application app = uiapp.Application;
            Document doc = uidoc.Document;

            TaskDialog.Show("Debug", "Hello There");
            
            // Get Excel Data - Currently selecting one file twice for this
            // Need to change the method to extract workbook instead of each worksheet
            // Maybe (or maybe not) add an overload that gets the WS if name is there and WB if not
            //List<string[]> wsSets = Util.GetExcelWS("Furniture Sets");
            //wsSets.RemoveAt(0);
            //List<string[]> wsTypes = Util.GetExcelWS("Furniture types");
            //wsTypes.RemoveAt(0);

            // Get Excel Workbook
            List<Spreadsheet> wsBook = Util.GetExcelWS();
            List<string[]> wsSets = new List<string[]>(); List<string[]> wsTypes = new List<string[]>();
            
            foreach (Spreadsheet ws in wsBook)
            {
                switch (ws.Name)
                {
                    case "Furniture sets":
                        wsSets = ws.Rows;
                        break;
                    case "Furniture types":
                        wsTypes = ws.Rows;
                        break;
                    default:
                        break;
                }
            }
            wsSets.RemoveAt(0); wsTypes.RemoveAt(0);




            //Build Furniture List from sets
            List<FurnitureSet> setList = new List<FurnitureSet>();

            foreach (string[] wsRow in wsSets)
            {
                //FurnitureSet curSet = new FurnitureSet(wsRow);
                // Wouldn't let me plug wsRow directly into the FurnitureSet so I indexed
                // Consider wsRow[i].Trim to clean up data
                FurnitureSet curSet = new FurnitureSet(wsRow[0], wsRow[1],wsRow[2]);
                setList.Add(curSet);
            }
            // Ignore this bit.  Use setList to look up Furniture Sets
            //FurnitureList furnList = new FurnitureList(setList);

            //Build Furniture Types from FurnType
            List<FurnitureType> typeList = new List<FurnitureType>();

            foreach (string[] wsRow in wsTypes)
            {
                FurnitureType curType = new FurnitureType(doc, wsRow[0], wsRow[1], wsRow[2]);
                typeList.Add(curType);
            }

            List<SpatialElement> roomList = Util.GetAllRooms(doc);
            
            using (Transaction t1 = new Transaction(doc))
            {
                t1.Start("Move in");
                
                foreach (SpatialElement room in roomList)
                {
                    string curSetID = Util.GetParameter(room, "Furniture Set");
                    FurnitureSet curSet = GetFurnSetByAlias(setList, curSetID);
                    
                    if (curSet != null)
                    {
                        foreach (string alias in curSet.FurnSet)
                        {
                            try
                            {
                                CreateFIinRoom(doc, room as Room, GetFSbyAlias(typeList, alias));
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }

                        }
                    }

                }
                t1.Commit();
            }


            // Furniture Count
            //List<SpatialElement> roomList = Util.GetAllRooms(doc);

            using (Transaction t2 = new Transaction(doc))
            {
                t2.Start("Count Furniture");

                foreach (SpatialElement room in roomList)
                {
                    List<FamilyInstance> curRoomFurn = Util.GetFurnInRoom(doc, room as Room);
                    int furnCount = curRoomFurn.Count;
                    // Set the Parameter
                    double intValue = Convert.ToDouble(furnCount);
                    Util.SetParameter(room, "Furniture Count", intValue);

                }

                t2.Commit();
            }
            
            return Result.Succeeded;
        }

        private void CreateFIinRoom(Document doc, Room room, FamilySymbol curFS)
        {
            LocationPoint roomLocation = room.Location as LocationPoint;
            XYZ roomPoint = roomLocation.Point;
            // Maybe add an instanced vector to room point, create at insertPoint so furniture is not stacked
            if (curFS != null)
                curFS.Activate();
            doc.Create.NewFamilyInstance(roomPoint, curFS, StructuralType.NonStructural);

        }

        private FamilySymbol GetFSbyAlias(List<FurnitureType> wsData, string alias)
        {
            // Getting Family Symbol in Class means I can remove the collector from this method
            //FilteredElementCollector coll = new FilteredElementCollector(doc)
            //    .OfClass(typeof(Family));
            //FamilySymbol curFS = null;

            foreach (FurnitureType furnType in wsData)
            {
                if (furnType.Name == alias)
                    return furnType.FamSymbol;
                // And I can just return the matching FamSymbol
                //{
                //    //curFS = Util.GetFamTypeByName(doc, furnType.FamName, furnType.FamType);
                //    return curFS;
                //}
            }

            return null;
        }

        private FurnitureSet GetFurnSetByAlias(List<FurnitureSet> sets, string id)
        {
            foreach (FurnitureSet set in sets)
            {
                if(set.SetName == id)
                    return set;
            }

            return null;
        }


    }
}
