#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
//using System.Linq;
using Forms = System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
//using Autodesk.Revit.DB.Architecture;
//using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.DB.Electrical;
//using Autodesk.Revit.DB.Mechanical;
//using Autodesk.Revit.DB.Plumbing;

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class CmdLoadGroups : IExternalCommand
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
            
            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Revit File";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Revit Files (*.rvt)|*.rvt";

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            string revitFile = ofd.FileName;

            // Any way to open and activate silently?
            UIDocument newUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            Document newDoc = newUIDoc.Document;

            FilteredElementCollector coll = new FilteredElementCollector(newDoc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsNotElementType();

            List<ElementId> groupIDList = new List<ElementId>();

            foreach (Element e in coll)
            {
                groupIDList.Add(e.Id);
            }

            Transform transform = null;
            CopyPasteOptions options = new CopyPasteOptions();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("CopyGroups");
                ElementTransformUtils.CopyElements(newDoc, groupIDList, doc, transform, options);
                t.Commit();
            }

            try
            {
                uiapp.OpenAndActivateDocument(doc.PathName);
                newUIDoc.SaveAndClose();
            }
            catch (Exception)
            {}

            TaskDialog.Show("Complete", "Loaded " + groupIDList.Count.ToString() + " groups into the current model.");

            return Result.Succeeded;
        }
    }
}
