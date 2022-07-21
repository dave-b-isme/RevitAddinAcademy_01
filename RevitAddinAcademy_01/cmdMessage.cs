#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class cmdMessage : IExternalCommand
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

            
            // interacting with app through ui
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            Level curLevel = GetLevelByName(doc, "Level 1");

            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Make Stuff");

                foreach (Element e in pickList)
                {
                    // is compares type, == compares value
                    if (e is CurveElement)
                    {
                        // e is Element, need to refer to it as a Curve Element to do curve stuff
                        CurveElement curve = (CurveElement)e;
                        //same thing as
                        CurveElement curve2 = e as CurveElement;

                        curveList.Add(curve);

                        // get some line properties
                        // get the graphicstyle from curve with AS GraphicStyle
                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                        // Get the geometry of the curve
                        Curve curCurve = curve.GeometryCurve;
                        XYZ startPoint = curCurve.GetEndPoint(0);
                        XYZ endPoint = curCurve.GetEndPoint(1);

                        //switch statement
                        // just does stuff, no need to collect and use if else
                        // no conditional logic in switch statement
                        switch (curGS.Name)
                        {
                            case "<Medium>":
                                Debug.Print("found a medium line");
                                break;

                            case "<Thin lines>":
                                Debug.Print("found a thin line");
                                break;

                            case "<Wide Lines":
                                Debug.Print("found a wide line");
                                break;

                            default:
                                Debug.Print("Found something else");
                                break;

                        }


                        

                        Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);

                        //Pipe newPipe = Pipe.Create(
                        //    doc,
                        //    curSystemType.Id,
                        //    curPipeType.Id,
                        //    curLevel.Id,
                        //    startPoint,
                        //    endPoint);


                        Debug.Print(curGS.Name);
                    }

                }
                t.Commit();

            }

            


            TaskDialog.Show("complete", curveList.Count.ToString());

            return Result.Succeeded;

        }

        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element curElem in collector)
            {
                WallType wallType = curElem as WallType;

                if (wallType.Name == wallTypeName)
                    return wallType;

            }
            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level level = curElem as Level;

                if (level.Name == levelName)
                    return level;

            }
            return null;
        }

        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in collector)
            {
                MEPSystemType curType = curElem as MEPSystemType;

                if (curType.Name == typeName)
                    return curType;

            }
            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curElem in collector)
            {
                PipeType curType = curElem as PipeType;

                if (curType.Name == typeName)
                    return curType;

            }
            return null;
        }

    }
}
