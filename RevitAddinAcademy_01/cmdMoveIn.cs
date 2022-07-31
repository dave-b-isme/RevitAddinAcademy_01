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
//using Autodesk.Revit.DB.Structure;
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
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            TaskDialog.Show("Debug", "Hello There");

            // Get Excel Data - Currently selecting one file twice for this
            // Need to change the method to extract workbook instead of each worksheet
            // Maybe (or maybe not) add an overload that gets the WS if name is there and WB if not
            List<string[]> wsSets = Util.GetExcelWS("Furniture Sets");
            wsSets.RemoveAt(0);
            List<string[]> wsTypes = Util.GetExcelWS("Furniture types");
            wsTypes.RemoveAt(0);

            //Build Furniture List from sets
            // I wonder if I can (or should) build the Furn List and Types directly in the Class?
            List<FurnitureSet> setList = new List<FurnitureSet>();

            foreach (string[] wsRow in wsSets)
            {
                //FurnitureSet curSet = new FurnitureSet(wsRow);
                // Wouldn't let me plug wsRow directly into the FurnitureSet so I indexed
                FurnitureSet curSet = new FurnitureSet(wsRow[0], wsRow[1],wsRow[2]);
                setList.Add(curSet);
            }

            FurnitureList furnList = new FurnitureList(setList);

            //Build Furniture Types from FurnType
            List<FurnitureType> typeList = new List<FurnitureType>();

            foreach (string[] wsRow in wsTypes)
            {
                FurnitureType curType = new FurnitureType(wsRow[0], wsRow[1], wsRow[2]);
                typeList.Add(curType);
            }
            FurnitureTypes furnTypes = new FurnitureTypes(typeList);
            int typeCount = furnTypes.GetFurnitureCount();


            // OK now what?




            // Furniture Count
            List<SpatialElement> roomList = Util.GetAllRooms(doc);

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

    }
}
