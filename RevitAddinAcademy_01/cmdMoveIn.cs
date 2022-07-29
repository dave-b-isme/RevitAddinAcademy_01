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

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            //dialog.InitialDirectory = @"C:\Users\David\Documents";
            dialog.InitialDirectory = @"%USERPROFILE%\Documents";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls; *.xlsx; *.xlsm | All Files | *.*";

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return Result.Failed;
            }

            string filePath = dialog.FileName;

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
