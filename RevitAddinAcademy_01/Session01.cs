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
    public class Session01 : IExternalCommand
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

            //int number = 1;
            //int x = 3;
            //int y = 5;
            //double testnum = 1.5;
            //string fizz = "FIZZ";
            //string buzz = "BUZZ";
            //string fizzbuzz = "FIZZBUZZ";
            //XYZ point = new XYZ(0, 0, 0);
            XYZ curPoint = new XYZ(0, 0, 0);

            List<string> strings = new List<string>();
            strings.Add("item 1");

            List<XYZ> points = new List<XYZ>();

            //int range = 100;
            //for (int i = 1; i <= range; i++)
            //{
            //    number = number + 1;
            //}

            string newString = "";
            foreach(string s in strings)
            {
                if(s == "item 1")
                {
                    newString = "go to 1";
                }
                else if(s == "item 2")
                {
                    newString = "go to 2";
                }
                else
                {
                    newString = "go to end";
                }

                newString = newString + s;
            }

            double newNumber = Method01(100, 100);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note");

            t.Start();

            TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "This is my text note", collector.FirstElementId());


            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

        internal double Method01(double a, double b)
        {
            double c = a + b;
            
            Debug.Print("Got here: " + c.ToString());

            return c;

        }
    }

}
