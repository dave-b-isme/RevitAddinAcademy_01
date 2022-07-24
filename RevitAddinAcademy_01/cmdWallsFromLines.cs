#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
//using Forms = System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class cmdWallsFromLines : IExternalCommand
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

            // Add option to delete lines after wall create
            // Add option to undo and recreate if wall needs flipping
            //

            List<string> wallTypes = GetAllWallTypeNames(doc);
            List<string> lineStyles = GetAllLineStyleNames(doc);

            frmWallsFromLines curForm = new frmWallsFromLines(wallTypes, lineStyles);
            curForm.Height = 300;
            curForm.Width = 360;
            curForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            if(curForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string selectedWallType = curForm.GetSelectedWallType();
                string selectedLineStyle = curForm.GetSelectedLineStyle();
                double wallHeight = curForm.GetWallHeight();
                bool isStructural = curForm.IsStructuralWall();

                WallType curWT = GetWallTypeByName(doc, selectedWallType);
                List<CurveElement> lineList = GetLinesByStyle(doc, selectedLineStyle);
                Level curLevel = GetLevelFromView(doc);

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Walls from Lines");

                    foreach(CurveElement line in lineList)
                    {
                        Curve curLine = line.GeometryCurve;
                        Wall newWall = Wall.Create(
                            doc,
                            curLine,
                            curWT.Id,
                            curLevel.Id,
                            wallHeight,
                            0,
                            false,
                            isStructural);
                    }
                    t.Commit();
                }

            }



            return Result.Succeeded;
        }

        private Level GetLevelFromView(Document doc)
        {
            View curView = doc.ActiveView;

            SketchPlane curSP = curView.SketchPlane;

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .WhereElementIsNotElementType();

            foreach(Level curLevel in collector)
            {
                if(curLevel.Name == curSP.Name)
                    return curLevel;
            }

            return collector.FirstElement() as Level;
        }

        private List<CurveElement> GetLinesByStyle(Document doc, string selectedLineStyle)
        {
            List<CurveElement> results = new List<CurveElement>();
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(CurveElement));

            foreach (CurveElement e in collector)
            {
                GraphicsStyle curGS = e.LineStyle as GraphicsStyle;

                if (curGS.Name == selectedLineStyle)
                {
                    results.Add(e);
                }

            }
            return results;
        }

        private WallType GetWallTypeByName(Document doc, string selectedWallType)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType));

            foreach (WallType e in collector)
            {
                if (e.Name == selectedWallType)
                    return e;
            }
            return null;
        }

        private List<string> GetAllLineStyleNames(Document doc)
        {
            List<string> results = new List<string>();
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfClass(typeof(CurveElement));

            foreach(CurveElement e in collector)
            {
                GraphicsStyle curGS = e.LineStyle as GraphicsStyle;

                if(results.Contains(curGS.Name) == false)
                {
                    results.Add(curGS.Name);
                }
            }
            return results;
        }

        private List<string> GetAllWallTypeNames(Document doc)
        {
            List<string> results = new List<string>();
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType));

            foreach (WallType e in collector)
            {
                results.Add(e.Name);
            }
            return results;
        }
    }
}
