using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.IO;

namespace RF_Reconstruction
{
    /// <summary>
    /// Summary description for RF_Reconstruction
    /// </summary>
    [CodedUITest]
    public class RF_Reconstruction
    {
        public RF_Reconstruction()
        {
        }

        //static variables
        static bool ExportFromImageArea;
        static String strImageLabelInUse;

        //static functions
        static void EmptyFolder(DirectoryInfo directoryInfo)
        {
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                file.Delete();
            }

            foreach (DirectoryInfo subfolder in directoryInfo.GetDirectories())
            {
                EmptyFolder(subfolder);
            }
        }

        //Use ClassInitialize to run code once before all tests 
        [ClassInitialize()]
        public static void MyTestInitialize(TestContext context)
        {
            // Clean up the destination folders
Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\Export\"), File.Delete);
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\B-Mode\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\B-Mode\Results\"));
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\Color\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\Color\Results\"));
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\M-Mode\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\M-Mode\Results\"));
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\NLC\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\NLC\Results\"));
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\Power\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\Power\Results\"));
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\RF Reconstruction\PW Doppler & PW Tissue Doppler\Exports\"), File.Delete);
            EmptyFolder(new DirectoryInfo(@"C:\Users\jqi2\Desktop\RF Reconstruction\PW Doppler & PW Tissue Doppler\Results\"));

            // Launch VevoLAB
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            proc.StartInfo.FileName = "C:\\Program Files\\VevoLAB\\Version-2.1.0\\VsiApp.exe";
            proc.Start();

            // Initialize static variables
            ExportFromImageArea = true;
        }

        //Use ClassCleanup to run code once after all tests have run
        [ClassCleanup()]
        public static void MyTestCleanup()
        {
            // Close VevoLAB
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            proc.StartInfo.FileName = "C:\\Program Files\\VevoLAB\\Version-2.1.0\\VsiApp.exe";
            proc.Kill();
        }

        [DataSource("System.Data.Odbc", "Dsn=Excel Files; Driver={Microsoft Excel Driver (*.xlsx)}; dbq=RF_Reconstruction.xlsx; defaultdir=|datadirectory|; driverid=790; maxbuffersize=2048; pagetimeout=5; readonly=true", "RFReconstruction$", DataAccessMethod.Sequential), DeploymentItem("RF_Reconstruction.xlsx"), TestMethod]
        public void RFReconstruction()
        {
            //First Level Hierarchy
            WinWindow VevoLAB = new WinWindow();
            VevoLAB.SearchProperties[WinWindow.PropertyNames.ClassName] = "VsiMain";

            WinWindow ImageExportReportWindow = new WinWindow();
            ImageExportReportWindow.SearchProperties[WinWindow.PropertyNames.Name] = "Image Export Report";

            //Second Level Hierarchy
            WinListItem StudyName = new WinListItem(VevoLAB);
            String strStudy = TestContext.DataRow["Study"].ToString();
            StudyName.SearchProperties[WinListItem.PropertyNames.Name] = strStudy;

            WinListItem SeriesName = new WinListItem(VevoLAB);
            String strSeries = TestContext.DataRow["Series"].ToString();
            SeriesName.SearchProperties[WinListItem.PropertyNames.Name] = strSeries;

            WinListItem ImageName = new WinListItem(VevoLAB);
            String strImageLabel = TestContext.DataRow["Image Label"].ToString();
            ImageName.SearchProperties[WinListItem.PropertyNames.Name] = strImageLabel;

            WinButton ExportButtonInStudyBrowser = new WinButton(VevoLAB);
            ExportButtonInStudyBrowser.SearchProperties[WinButton.PropertyNames.Name] = "Export";

            WinButton PauseButton = new WinButton(VevoLAB);
            PauseButton.SearchProperties[WinButton.PropertyNames.Name] = "Forw";

            WinButton FirstFrameButton = new WinButton(VevoLAB);
            FirstFrameButton.SearchProperties[WinButton.PropertyNames.Name] = "Home";

            WinButton ExportButtonInModeWindow = new WinButton(VevoLAB);
            ExportButtonInModeWindow.SearchProperties[WinButton.PropertyNames.Name] = "Export";

            WinButton BacktoStudyBrowserButton = new WinButton(VevoLAB);
            BacktoStudyBrowserButton.SearchProperties[WinButton.PropertyNames.Name] = "S";

            WinTreeItem DesktopTreeItemInExportImageWindow = new WinTreeItem(VevoLAB);
            DesktopTreeItemInExportImageWindow.SearchProperties[WinTreeItem.PropertyNames.Name] = "Desktop";
            WinTreeItem ExportTreeItemInExportImageWindow = new WinTreeItem(DesktopTreeItemInExportImageWindow);
ExportTreeItemInExportImageWindow.SearchProperties[WinTreeItem.PropertyNames.Name] = "Export";

            WinRadioButton ExportTypeRadioButton = new WinRadioButton(VevoLAB);
            String strExportType = TestContext.DataRow["Export Type"].ToString();
            ExportTypeRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = strExportType;

            WinEdit SaveAsInExportImageWindow = new WinEdit(VevoLAB);
            SaveAsInExportImageWindow.SearchProperties[WinEdit.PropertyNames.Name] = "Save As";

            String strFileType = TestContext.DataRow["File Type"].ToString();
            WinComboBox FileTypeComboBox = new WinComboBox(VevoLAB);

            WinButton OKButtonInExportImageWindow = new WinButton(VevoLAB);
            OKButtonInExportImageWindow.SearchProperties[WinButton.PropertyNames.Name] = "OK";

            WinButton OKButtonInImageExportReportWindow = new WinButton(ImageExportReportWindow);
            OKButtonInImageExportReportWindow.SearchProperties[WinButton.PropertyNames.Name] = "OK";

            // If it is a new image
            if (strImageLabel != "")
            {
                // If not the first image, Switch to 'Study Browser' window
                if (BacktoStudyBrowserButton.Exists)
                {
                    // Click Back to Study Browser button
                    Mouse.Click(BacktoStudyBrowserButton, BacktoStudyBrowserButton.GetClickablePoint());
                }

                // Switch to 'Mode' window by loading a image
                //// Double-Click to expand the Study in the Study Browser
                if (strStudy != "")
                {
                    Mouse.DoubleClick(StudyName, StudyName.GetClickablePoint());
                }

                //// Double-Click to expand the Series in the Study Browser
                if (strSeries != "")
                {
                    Mouse.DoubleClick(SeriesName, SeriesName.GetClickablePoint());
                }

                //// Double-Click the Image to load
                Mouse.DoubleClick(ImageName, ImageName.GetClickablePoint());
                
                ////// Pause playback
                Mouse.Click(PauseButton, PauseButton.GetClickablePoint());

                ////// Move to the first frame
                Mouse.Click(FirstFrameButton, FirstFrameButton.GetClickablePoint());                
            }

            // Swith to 'Export Image' window to set parameters and export
            if (TestContext.DataRow["Export Button"].ToString() == "Image Area")
            {
                if (ExportFromImageArea == false && strImageLabel == "")
                {
                    //// Double-Click the Image to load
                    WinListItem ImageNameInUse = new WinListItem(VevoLAB);
                    ImageNameInUse.SearchProperties[WinListItem.PropertyNames.Name] = strImageLabelInUse;
                    Mouse.DoubleClick(ImageNameInUse, ImageNameInUse.GetClickablePoint());
                }                

                //// Export from Mode Window
                Mouse.Click(ExportButtonInModeWindow, ExportButtonInModeWindow.GetClickablePoint());

                ExportFromImageArea = true;

                if (strImageLabel != "")
                {
                    strImageLabelInUse = strImageLabel;
                }
            }
            else if (TestContext.DataRow["Export Button"].ToString() == "Study Browser")
            {
                //// Click Back to Study Browser button
                Mouse.Click(BacktoStudyBrowserButton, BacktoStudyBrowserButton.GetClickablePoint());

                //// Export from Study Browser
                Mouse.Click(ExportButtonInStudyBrowser, ExportButtonInStudyBrowser.GetClickablePoint());

                ExportFromImageArea = false;

                if (strImageLabel != "")
                {
                    strImageLabelInUse = strImageLabel;
                }
            }
            else
            {
                if (ExportFromImageArea == true)
                {
                    //// Export from Mode Window
                    Mouse.Click(ExportButtonInModeWindow, ExportButtonInModeWindow.GetClickablePoint());
                }
                else if (ExportFromImageArea == false)
                {
                    //// Export from Study Browser
                    Mouse.Click(ExportButtonInStudyBrowser, ExportButtonInStudyBrowser.GetClickablePoint());
                }
            }

            //// Click 'Desktop' -> 'RF Reconstruction' tree item as the export destination
            Mouse.Click(ExportTreeItemInExportImageWindow, ExportTreeItemInExportImageWindow.GetClickablePoint());

            //// If it is a new 'Export Type'
            if (strExportType != "")
            {
                //// Click one of the 'Export Type' radio button
                Mouse.Click(ExportTypeRadioButton, ExportTypeRadioButton.GetClickablePoint());
            }

            //// Give a name in 'Save As' text box
            int TestResultRow = TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow);
            int ExcelRow = TestResultRow + 2;
            String TestCaseNumber = TestContext.DataRow["Save As"].ToString();
SaveAsInExportImageWindow.Text = string.Format("ExcelRow_{0}_{1}", ExcelRow, TestCaseNumber);

            //// Select a 'File Type' from the combo box            
            FileTypeComboBox.SelectedItem = strFileType;

            // Click 'OK' button in 'Export Image' window
            Mouse.Click(OKButtonInExportImageWindow, OKButtonInExportImageWindow.GetClickablePoint());

            // Click 'OK' button in 'Image Export Report' pop-up window
            Mouse.Click(OKButtonInImageExportReportWindow, OKButtonInImageExportReportWindow.GetClickablePoint());
        }

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.s
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;
    }
}
