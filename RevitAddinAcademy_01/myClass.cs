using System;
//ToList command in a collections statement somewhere, conflict
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
//using Forms = System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Architecture;


namespace RevitAddinAcademy_01
{
    internal class myClass
    {
    }

    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<string> FavColors { get; set; }

        public Employee(string name, int age, string favColors)
        {
            Name = name;
            Age = age;
            FavColors = FormatColorList(favColors);
        }

        private List<string> FormatColorList(string colorList)
        {
            // formats comma delineated text into list
            List<string> returnList = colorList.Split(',').ToList();
            return returnList;
        }
    }

    public class Employees
    {
        public List<Employee> EmployeeList { get; set; }

        public Employees(List<Employee> employees)
        {
            EmployeeList = employees;
        }

        public int GetEmployeeCount()
        {
            return EmployeeList.Count;
        }

    }

    // How static is a static class?  Can it return variable info from an external file?
    public static class Utilities
    {
        public static string GetTextFromClass()
        {
            return "I got this text from a static class";
        }

        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms);

            List<SpatialElement> roomList = new List<SpatialElement>();

            foreach(Element curElem in collector)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                roomList.Add(curRoom);
            }
            return roomList;
        }

        public static FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Family));

            foreach(Element e in collector)
            {
                Family curFamily = e as Family;

                if(curFamily.Name == familyName)
                {
                    ISet<ElementId> famSymbolIdList = curFamily.GetFamilySymbolIds();
                    foreach(ElementId famSymbolId in famSymbolIdList)
                    {
                        FamilySymbol curFS = doc.GetElement(famSymbolId) as FamilySymbol;
                        if(curFS.Name == typeName)
                            return curFS;
                    }
                }
            }
            return null;
        }

        public static string GetParamValueAsString(Element curElem, string paramName)
        {
            foreach (Parameter CurParam in curElem.Parameters)
            {
                if(CurParam.Definition.Name == paramName)
                    return CurParam.AsString();
            }

            return null;
        }

        public static double GetParamValueAsDouble(Element curElem, string paramName)
        {
            foreach (Parameter CurParam in curElem.Parameters)
            {
                if (CurParam.Definition.Name == paramName)
                    return CurParam.AsDouble();
            }

            return 0;
        }
        public static void SetParamValue(Element curElem, string paramName, string paramValue)
        {
            foreach(Parameter CurParam in curElem.Parameters)
            {
                if (CurParam.Definition.Name == paramName)
                    CurParam.Set(paramValue);
            }

        }
        //OVERLOAD method has two input types, one for string another for double
        // maybe pick up all those types by name as OVERLOAD methods
        public static void SetParamValue(Element curElem, string paramName, double paramValue)
        {
            foreach (Parameter CurParam in curElem.Parameters)
            {
                if (CurParam.Definition.Name == paramName)
                    CurParam.Set(paramValue);
            }

        }

    }
}
