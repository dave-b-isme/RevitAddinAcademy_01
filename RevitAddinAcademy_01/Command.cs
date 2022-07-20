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

#endregion

namespace RevitAddinAcademy_01
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            //Comment Line
            //Instatiate a dialog box
            //FileBrowserDialog is another Option - folders vs files maybe
            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            //dialog.Multiselect = false;
            dialog.Multiselect = true;
            dialog.Filter = "Revit Files | *.rvt; *.rfa";

            //string filePath = "";
            // Hmmm. No need to set a value when instantiating an array
            string[] filePaths;

            if(dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                //filePath = dialog.FileName;
                filePaths = dialog.FileNames;
            }

            Forms.FolderBrowserDialog folderDialog = new Forms.FolderBrowserDialog();

            string folderPath = "";
            if(folderDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                folderPath = folderDialog.SelectedPath;
            }

            Tuple<string, int> t1 = new Tuple<string, int>("string 1", 55);
            Tuple<string, int> t2 = new Tuple<string, int>("string 2", 155);

            TestStruct struct1;
            struct1.Name = "Structure 1";
            struct1.Value = 100;
            struct1.Value2 = 10.5;

            //Use structures to collect data in organized way
            TestStruct struct2 = new TestStruct("structure 1", 10, 1004.4);

            //double var = struct2.addNumber();

            double value3 = struct1.Value + struct2.Value2;

            List<TestStruct> structList = new List<TestStruct>();
            structList.Add(struct1);

            Level newLevel = Level.Create(doc, 100);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType curVFT = null;
            ViewFamilyType curRCPVFT = null;

            //foreach(Element curElem in collector) Run ops on specific class and not on gen element sometimes
            foreach(ViewFamilyType curElem in collector)
            {
                //if(curElem.Name == "Floor Plan")..
                if (curElem.ViewFamily == ViewFamily.FloorPlan)
                {
                    curVFT = curElem;
                }
                
                else
                    if(curElem.ViewFamily == ViewFamily.CeilingPlan)
                {
                    curRCPVFT = curElem;
                }



            }

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector2.WhereElementIsElementType();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Making Plans");
                
                ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, newLevel.Id);
                curRCP.Name = curRCP.Name + "RCP";

                View existingView = GetViewByName(doc, "Level 1");

                ViewSheet newSheet = ViewSheet.Create(doc, collector2.FirstElementId());

                if(existingView != null)
                {
                    //Don't do it
                }
                else
                {
                    // Do IT
                }

                Viewport newVP = Viewport.Create(doc, newSheet.Id, curPlan.Id, new XYZ(0, 0, 0));

                newSheet.Name = "TEST SHEET";
                newSheet.SheetNumber = "A10010101";

                
                foreach(Parameter curParam in newSheet.Parameters)
                {
                    if(curParam.Definition.Name == "Drawn By")
                    {
                        curParam.Set("DB");                        
                    }
                }


                t.Commit();

            }
                



            return Result.Succeeded;
        }



        // METHOD Get View By Name
        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach(View curView in collector)
            {
                if(curView.Name == viewName)
                {
                    return curView;
                }
            }

            return null;

        }

        //Structure
        internal struct TestStruct
        {
            public string Name;
            public int Value;
            public double Value2;

            //Constructor
            public TestStruct(string name, int value, double value2)
            {
                Name = name;
                Value = value;
                Value2 = value2;
            }

            public double addNumber()
            {
                return Value + Value2;
            }

        }
    }
}
