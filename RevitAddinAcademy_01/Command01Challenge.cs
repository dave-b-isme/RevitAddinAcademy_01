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

            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create Text Note");
            t.Start();
            
            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                string curNote = "";
                if (i % 3 == 0)
                {
                    curNote = curNote + "FIZZ";                    
                }
                if (i % 5 == 0)
                {
                    curNote += "BUZZ";
                }
                else if (i % 3 != 0)
                {
                    curNote = i.ToString();
                }

                //TextNote printNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, curNote, collector.FirstElementId());
                TextNote.Create(doc, doc.ActiveView.Id, curPoint, curNote, collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);

            }


            t.Commit();
            t.Dispose();


            return Result.Succeeded;
        }



    }
}
