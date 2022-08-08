#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

#endregion

namespace RevitAddinAcademy_01
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            // Create Ribbon Tab
            try
            {
                a.CreateRibbonTab("RAA Tab");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            // Create Ribbon Panel
            RibbonPanel curPanel = CreateRibbonPanel(a, "RAA Tab", "RAA Panel");

            // Create Data for Buttons
            //PushButtonData pData0 = new PushButtonData("button0", "Button 0", GetAssemblyName(), "RevitAddinAcademy_01.Command");
            //pData0.Image = BitmapToImageSource(Properties.Resources.Blue_16);
            //pData0.LargeImage = BitmapToImageSource(Properties.Resources.Blue_32);
            ButtonData bData1 = new ButtonData("tool1", "PopUp \nTest",
                Properties.Resources._1_16,
                Properties.Resources._1_32,
                "Command", "Hello There");
            ButtonData bData2 = new ButtonData("tool2", "Delete \nBackups",
                Properties.Resources._2_16,
                Properties.Resources._2_32,
                "CmdDeleteBackups", "Delete Backup files from selected directory");
            CreatePushButton(curPanel, bData1);
            CreatePushButton(curPanel, bData2);

            ButtonData bData3 = new ButtonData("tool3", "Move In",
                Properties.Resources._3_16,
                Properties.Resources._3_32,
                "CmdMoveIn", "Move Furniture into Rooms");
            ButtonData bData4 = new ButtonData("tool4", "Load \nGroups",
                Properties.Resources._4_16,
                Properties.Resources._4_32,
                "CmdLoadGroups", "Load Groups from Revit File");
            ButtonData bData5 = new ButtonData("tool5", "Create \nWalls",
                Properties.Resources._5_16,
                Properties.Resources._5_32,
                "CmdWallsFromLines", "Create Walls from Lines");
            List<ButtonData> stackList = new List<ButtonData>();
            stackList.Add(bData3); stackList.Add(bData4); stackList.Add(bData5);
            CreateStack(curPanel, stackList);

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");
            ButtonData bData6 = new ButtonData("tool6", "Load \nGroups",
                Properties.Resources._6_16,
                Properties.Resources._6_32,
                "CmdLoadGroups", "Load Groups from Revit File");
            ButtonData bData7 = new ButtonData("tool7", "Create \nWalls",
                Properties.Resources._7_16,
                Properties.Resources._7_32,
                "CmdWallsFromLines", "Create Walls from Lines");
            List<ButtonData> splitList = new List<ButtonData>();
            splitList.Add(bData6); splitList.Add(bData7);
            CreateSplit(curPanel, sData1, splitList);
            //SplitButton sb1 = curPanel.AddItem(sData1) as SplitButton;
            //sb1.AddPushButton(bData6); sb1.AddPushButton(bData7);


            PulldownButtonData pdbData = new PulldownButtonData("pdButton1", "MoreTools");
            pdbData.Image = BitmapToImageSource(Properties.Resources.Blue_16);
            pdbData.LargeImage = BitmapToImageSource(Properties.Resources.Blue_32);
            ButtonData bData8 = new ButtonData("tool8", "PopUp \nTest",
                Properties.Resources._8_16,
                Properties.Resources._8_32,
                "Command", "Hello There");
            ButtonData bData9 = new ButtonData("tool9", "PopUp \nTest",
                Properties.Resources._9_16,
                Properties.Resources._9_32,
                "Command", "Hello There");
            ButtonData bData10 = new ButtonData("tool10", "PopUp \nTest",
                Properties.Resources.Blue_16,
                Properties.Resources.Blue_32,
                "Command", "Hello There");
            List<ButtonData> pdb = new List<ButtonData>();
            pdb.Add(bData8); pdb.Add(bData9); pdb.Add(bData10);
            CreatePullDown(curPanel, pdbData, pdb);


            return Result.Succeeded;
        }

        private void CreateStack(RibbonPanel curPanel, List<ButtonData> stackList)
        {
            List<PushButtonData> dataList = new List<PushButtonData>();

            foreach (ButtonData bData in stackList)
            {
                PushButtonData curData = new PushButtonData(bData.Name, bData.Text, GetAssemblyName(), bData.Command);
                curData.Image = bData.Icon;
                curData.LargeImage = bData.LargeIcon;
                dataList.Add(curData);
            }

            curPanel.AddStackedItems(dataList[0], dataList[1], dataList[2]);
        }

        private void CreatePullDown(RibbonPanel curPanel, PulldownButtonData pdbData, List<ButtonData> bList)
        {
            List<PushButtonData> dataList = new List<PushButtonData>();
            PulldownButton pdb = curPanel.AddItem(pdbData) as PulldownButton;
            
            foreach (ButtonData bData in bList)
            {
                PushButtonData curData = new PushButtonData(bData.Name, bData.Text, GetAssemblyName(), bData.Command);
                curData.Image = bData.Icon;
                curData.LargeImage = bData.LargeIcon;
                dataList.Add(curData);
            }
            foreach (PushButtonData curData in dataList)
            {
                pdb.AddPushButton(curData);
            }
        }

        private void CreateSplit(RibbonPanel curPanel, SplitButtonData sbData, List<ButtonData> bList)
        {
            List<PushButtonData> dataList = new List<PushButtonData>();
            SplitButton curSplit = curPanel.AddItem(sbData) as SplitButton;
            
            foreach (ButtonData bData in bList)
            {
                PushButtonData curData = new PushButtonData(bData.Name, bData.Text, GetAssemblyName(), bData.Command);
                curData.Image = bData.Icon;
                curData.LargeImage = bData.LargeIcon;
                dataList.Add(curData);
            }
            foreach (PushButtonData curData in dataList)
            {
                curSplit.AddPushButton(curData);
            }
        }

        private void CreatePushButton(RibbonPanel curPanel, ButtonData bData)
        {
            //Take ButtonData and Create New PushButtonData
            PushButtonData curData = new PushButtonData(bData.Name, bData.Text, GetAssemblyName(), bData.Command);
            curData.Image = bData.Icon;
            curData.LargeImage = bData.LargeIcon;

            // Create PushButton from PushButtonData (separate method?)
            curPanel.AddItem(curData);
        }

        private BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                bm.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = ms;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }
        
        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach(RibbonPanel tmpPanel in a.GetRibbonPanels())
            {
                if(tmpPanel.Name == panelName)
                    return tmpPanel;
            }
            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);
            return returnPanel;
        }
        
        
        public Result OnShutdown(UIControlledApplication a)
        {
            

            return Result.Succeeded;
        }
    }
}
