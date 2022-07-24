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
    public class cmdElementsFromLines : IExternalCommand
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
            // uidoc.Selection collects IList for some reason
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select lines");


            // Lets try selecting only lines
            //ISelectionFilter lineFilter = new LineSelectionFilter();
            // OK nvm. Looks like I need to create a class to make that work. Not ready for that
            // Would probably make it easier to deal with those circles though.
            List<CurveElement> lineList = new List<CurveElement>();

            // I would like to use the same method for all of these
            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            WallType curGlazType = GetWallTypeByName(doc, "Storefront");
            Level curLevel = GetLevelByName(doc, "Level 1");

            MEPSystemType curPipeSysType = GetSystemTypeByName(doc, "Domestic Hot Water");
            MEPSystemType curDuctSysType = GetSystemTypeByName(doc, "Exhaust Air");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");
            DuctType curDuctType = GetDuctTypeByName(doc, "Default");

            int glaz = 0; int wall = 0; int pipe = 0; int duct = 0; int oth = 0;
            int curcles = 0;
            int counter = 0; int lineCount = 0; int arcCount = 0;

            // Count the curvy curves


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Reveal");
                // Brings everything into the loop as elements, creates lists within loop to sort them
                foreach (Element e in pickList)
                {
                    // is compares type, == compares value
                    if (e is CurveElement)
                    {
                        // e is Element, need to refer to it as a Curve Element to do curve stuff
                        CurveElement line = (CurveElement)e;
                        //same thing as
                        //CurveElement line2 = e as CurveElement;

                        lineList.Add(line);

                        // get some line properties
                        // get the graphicstyle from curve by taking the linestyle and getting it with AS GraphicStyle
                        GraphicsStyle curGS = line.LineStyle as GraphicsStyle;

                        // Get the geometry of the curve
                        Curve curLine = line.GeometryCurve;
                        XYZ startPoint = new XYZ(0, 0, 0);
                        XYZ endPoint = new XYZ(0, 0, 0);
                        
                        try
                        {
                            startPoint = curLine.GetEndPoint(0);
                            endPoint = curLine.GetEndPoint(1);
                        }
                        catch (Exception ep)
                        {
                            Debug.WriteLine(ep.ToString());
                            curcles++;                             
                        }


                        //switch statement
                        // just does stuff, no need to collect and use if else
                        // no conditional logic in switch statement

                        // CHALLENGE - Let's try making the stuff with switch statements
                        // QUESTION - can you put wildcard matches or starts with or regex stuff in switch cases?
                        // QUESTION - does the break in case stop the switch statement or does it keep evaluating other cases?
                        //      Perhaps one of these does it: You need to break;, throw, goto, or return from each of your case labels. In a loop you may also continue.

                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Debug.Print("found a Storefront line");
                                Wall newGlaz = Wall.Create(
                                    doc, 
                                    curLine, 
                                    curGlazType.Id, curLevel.Id, 
                                    20, 0, 
                                    false, false);
                                glaz++;
                                break;

                            case "A-WALL":
                                Debug.Print("found a Wall line");
                                Wall newWall = Wall.Create(
                                    doc,
                                    curLine,
                                    curWallType.Id, curLevel.Id,
                                    15, 0,
                                    false, false);
                                wall++;
                                break;

                            case "M-DUCT":
                                Debug.Print("found a Duct line");
                                Duct newDuct = Duct.Create(
                                    doc,
                                    curDuctSysType.Id, curDuctType.Id, curLevel.Id,
                                    startPoint, endPoint
                                    );
                                duct++;
                                break;

                            case "P-PIPE":
                                Debug.Print("found a Pipe line, should have called USA-DIG");
                                Pipe newPipe = Pipe.Create(
                                    doc,
                                    curPipeSysType.Id,
                                    curPipeType.Id,
                                    curLevel.Id,
                                    startPoint,
                                    endPoint);
                                pipe++;
                                break;

                            default:
                                Debug.Print("Found something else");
                                oth++;
                                break;

                        }
                        
                        Debug.Print(curGS.Name);
                    }

                    if (e is CurveElement)
                        counter++;
                    if (e is ModelLine)
                        lineCount++;
                    if (e is ModelArc)
                        arcCount++;

                }
                t.Commit();

            }


            Debug.Print(glaz.ToString());

            TaskDialog.Show("Complete", "I found "+ lineList.Count.ToString() + " Lines" + "\r\n" +
                "\r\n" +
                "I converted: " + "\r\n" +
                wall.ToString() + " Walls" + "\r\n" +
                glaz.ToString() + " Storefronts" + "\r\n" +
                duct.ToString() + " Ducts" + "\r\n" +
                pipe.ToString() + " Pipes" + "\r\n" +
                "\r\n" +
                "And I found " + oth.ToString() + " other lines" + "\r\n"
                );

            TaskDialog.Show("Complete", "I also found " + curcles.ToString() + " circles some trickster left in there" +"\r\n" +
                "\r\n" +
                "CurveElement Count: " + counter.ToString() + "\r\n" +
                "ModelLine Count: " + lineCount.ToString() + "\r\n" +
                "ModelArc Count: " + arcCount.ToString() + "\r\n"
                );

            return Result.Succeeded;

        }

        //// QUESTION Is it possible to add the class as an input and use one method to get any kind of type?
        ////
        //private Type? GetTypeByName(Document doc, Type? famType, string typeName)
        //{
        //    FilteredElementCollector collector = new FilteredElementCollector(doc);
        //    collector.OfClass();
        //    
        //    if (famType is DuctType)
        //
        //    foreach (Element curElem in collector)
        //    {
        //        DuctType curType = curElem as DuctType;
        //
        //        if (curType.Name == typeName)
        //            return curType;
        //    }
        //    return null;
        //}

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

        private DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element curElem in collector)
            {
                DuctType curType = curElem as DuctType;

                if (curType.Name == typeName)
                    return curType;

            }
            return null;
        }



    }
}
