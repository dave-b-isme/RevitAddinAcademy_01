#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class Command01Challenge : IExternalCommand
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

            //string text = "Revit Add-in Academy";
            //string fileName = doc.PathName;

            double offset = 0.05;
            double offsetCalc = offset * doc.ActiveView.Scale;

            XYZ curPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);

            string curNote = "curNote";

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note");
            t.Start();

            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                TextNote printNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, fizzBuzz(i), collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);
            }

            //int range = 100;
            //for (int i = 1; i <= range; i++)
            //{
                
            //    if(i % 3 == 0 & i % 5 == 0)
            //    {
            //        TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "FIZZBUZZ", collector.FirstElementId());
                    
            //    }
            //    else if(i % 3 == 0)
            //    {
            //        TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "FIZZ", collector.FirstElementId());
                    
            //    }
            //    else if(i % 5 == 0)
            //    {
            //        TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "BUZZ", collector.FirstElementId());
                    
            //    }
            //    else
            //    {
            //        TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "Number" + i.ToString(), collector.FirstElementId());
                    
            //    }
                    


                
            //    curPoint = curPoint.Subtract(offsetPoint);

            //}

            t.Commit();
            t.Dispose();


            return Result.Succeeded;
        }

        //internal double Method01(double a, double b)
        //{
        //    double c = a + b;

        //    Debug.Print("Got here: " + c.ToString());

        //    return c;
        //}

        internal string fizzBuzz(int i)
        {
            
            if (i % 3 == 0 && i % 5 == 0)
            {
                string curNote = "FIZZBUZZ";
            }
            else if (i % 3 == 0 || i % 5 == 0)
            {
                if (i % 3 == 0)
                {
                    string curNote = "FIZZ";
                }
                else
                {
                    string curNote = "BUZZ";
                }
            }
            else
            {
                string curNote = "";
            }

            return curNote;
            

        }

    }
}
