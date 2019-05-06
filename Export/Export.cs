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

namespace Export
{
    /// <summary>
    /// Summary description for Export
    /// </summary>
    [CodedUITest]
    public class Export
    {
        public Export()
        {
        }

        //static variables
        static int AppNo;
        static int ExcelNo;
        static bool ExportFromImageArea;
        static String strImageLabelInUse;

        //Use ClassInitialize to run code once before all tests 
        [ClassInitialize()]
        public static void MyTestInitialize(TestContext context)
        {
            // Clean up the destination folder
            Array.ForEach(Directory.GetFiles(@"C:\Users\jqi2\Desktop\Export\"), File.Delete);

            // Launch VevoLAB
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            proc.StartInfo.FileName = "C:\\Program Files\\VevoLAB\\Version-2.1.0\\VsiApp.exe";
            proc.Start();

            // Initialize static variables
            AppNo = 0;
            ExcelNo = 0;
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

        [DataSource("System.Data.Odbc", "Dsn=Excel Files; Driver={Microsoft Excel Driver (*.xlsx)}; dbq=Export.xlsx; defaultdir=|datadirectory|; driverid=790; maxbuffersize=2048; pagetimeout=5; readonly=true", "CheckExportFileType$", DataAccessMethod.Sequential), DeploymentItem("Export.xlsx"), TestMethod]
        public void CheckExportFileType()
        {

            //First Level Hierarchy
            WinWindow VevoLAB = new WinWindow();
            VevoLAB.SearchProperties[WinWindow.PropertyNames.ClassName] = "VsiMain";

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

            WinRadioButton ExportTypeRadioButton = new WinRadioButton(VevoLAB);
            String strExportType = TestContext.DataRow["Export Type"].ToString();
            ExportTypeRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = strExportType;

            WinComboBox FileTypeComboBox = new WinComboBox(VevoLAB);

            WinButton CancelButtonInExportImageWindow = new WinButton(VevoLAB);
            CancelButtonInExportImageWindow.SearchProperties[WinButton.PropertyNames.Name] = "Cancel";

            // If it is a new image
            if (strImageLabel != "")
            {
                //If not the first image
                if (CancelButtonInExportImageWindow.Exists)
                {
                    // Click 'Cancel' button in the Export Image Window to switch back to Sutdy Browser
                    Mouse.Click(CancelButtonInExportImageWindow, CancelButtonInExportImageWindow.GetClickablePoint());
                }

                // Double-Click to expand the Study in the Study Browser
                if (strStudy != "")
                {
                    Mouse.DoubleClick(StudyName, StudyName.GetClickablePoint());
                }

                // Double-Click to expand the Series in the Study Browser
                if (strSeries != "")
                {
                    Mouse.DoubleClick(SeriesName, SeriesName.GetClickablePoint());
                }

                // Click the Image to select
                Mouse.Click(ImageName, ImageName.GetClickablePoint());

                // Click 'Export' button
                Mouse.Click(ExportButtonInStudyBrowser, ExportButtonInStudyBrowser.GetClickablePoint());
            }

            // If it is a new 'Export Type'
            if (strExportType != "" || strExportType == "END")
            {
                // Compare the total number of 'File Types' from the Excel and the Combo Box
                try
                {
                    Assert.AreEqual(ExcelNo, AppNo, "the total number of 'File Types' from the Excel and the Combo Box are NOT equal");
                }
                catch
                {
                    if (strExportType != "END")
                    {
                        // Click one of the 'Export Type' radio button
                        Mouse.Click(ExportTypeRadioButton, ExportTypeRadioButton.GetClickablePoint());

                        // Count total number of 'File Types' in the File Type combo box
                        AppNo = FileTypeComboBox.Items.Count;

                        // Start to count the total number of 'File Types' in the Excel 
                        ExcelNo = 1;

                        // Select a 'File Type' from the combo box            
                        FileTypeComboBox.SelectedItem = TestContext.DataRow["File Type"].ToString();
                    }

                    //to force script to fail
                    // Note the Data Row Number n in the test results output is different than the Data Row Number m in the Excel
                    // m = n + 2
                    Assert.Fail("the total number of 'File Types' from the Excel and the Combo Box are NOT equal");
                }

                if (strExportType != "END")
                {
                    // Click one of the 'Export Type' radio button
                    Mouse.Click(ExportTypeRadioButton, ExportTypeRadioButton.GetClickablePoint());

                    // Count total number of 'File Types' in the File Type combo box
                    AppNo = FileTypeComboBox.Items.Count;

                    // Start to count the total number of 'File Types' in the Excel 
                    ExcelNo = 1;
                }
            }
            else
            {
                // Continue to count the total number of 'File Types' in the Excel 
                ExcelNo = ExcelNo + 1;
            }

            if (strExportType != "END")
            {
                // Select a 'File Type' from the combo box            
                FileTypeComboBox.SelectedItem = TestContext.DataRow["File Type"].ToString();
            }
            else
            {
                // Click 'Cancel' button in the Export Image Window to
                // Switch back to Study Browser for next test method
                Mouse.Click(CancelButtonInExportImageWindow, CancelButtonInExportImageWindow.GetClickablePoint());
            }
        }

        [DataSource("System.Data.Odbc", "Dsn=Excel Files; Driver={Microsoft Excel Driver (*.xlsx)}; dbq=Export.xlsx; defaultdir=|datadirectory|; driverid=790; maxbuffersize=2048; pagetimeout=5; readonly=true", "Exports$", DataAccessMethod.Sequential), DeploymentItem("Export.xlsx"), TestMethod]
        public void Exports()
        {
            //First Level Hierarchy
            WinWindow VevoLAB = new WinWindow();
            VevoLAB.SearchProperties[WinWindow.PropertyNames.ClassName] = "VsiMain";

            WinWindow NewMeasurementPackagWindow = new WinWindow();
            NewMeasurementPackagWindow.SearchProperties[WinWindow.PropertyNames.Name] = "New Measurement Package";

            WinWindow PackageExistsWindow = new WinWindow();
            PackageExistsWindow.SearchProperties[WinWindow.PropertyNames.Name] = "Package Exists";

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

            WinButton ImageProcessingButton = new WinButton(VevoLAB);
            ImageProcessingButton.SearchProperties[WinButton.PropertyNames.Name] = "PostProc";

            WinComboBox DisplayLayoutComboBox = new WinComboBox(VevoLAB);
            DisplayLayoutComboBox.SearchProperties[WinComboBox.PropertyNames.Name] = "Display Layout";

            WinCheckBox InvertCheckBox = new WinCheckBox(VevoLAB);
            InvertCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Invert";

            WinCheckBox RFOverlayCheckBox = new WinCheckBox(VevoLAB);
            RFOverlayCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "RF Overlay";

            WinButton PhysioButton = new WinButton(VevoLAB);
            PhysioButton.SearchProperties[WinButton.PropertyNames.Name] = "Physio";

            WinCheckBox ViewPhysiologyCheckBox = new WinCheckBox(VevoLAB);
            ViewPhysiologyCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "View Physiology";

            WinCheckBox ECGCheckBox = new WinCheckBox(VevoLAB);
            ECGCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "ECG";

            WinCheckBox RespirationCheckBox = new WinCheckBox(VevoLAB);
            RespirationCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Respiration";

            WinCheckBox InvertRespirationCheckBox = new WinCheckBox(VevoLAB);
            InvertRespirationCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Invert";

            WinCheckBox BPCheckBox = new WinCheckBox(VevoLAB);
            BPCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "BP";

            WinCheckBox BPDerivativeCheckBox = new WinCheckBox(VevoLAB);
            BPDerivativeCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "BP Derivative";

            WinCheckBox TemperatureCheckBox = new WinCheckBox(VevoLAB);
            TemperatureCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Temperature";

            WinButton PreferencesButton = new WinButton(VevoLAB);
            PreferencesButton.SearchProperties[WinButton.PropertyNames.Name] = "Prefs";

            WinTabPage GeneralTabPage = new WinTabPage(VevoLAB);
            GeneralTabPage.SearchProperties[WinTabPage.PropertyNames.Name] = "     General     ";

            WinCheckBox ShowDateTimeonImageHeaderCheckBox = new WinCheckBox(VevoLAB);
            ShowDateTimeonImageHeaderCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Show Date/Time on Image Header";

            WinTabPage MeasurementTabPage = new WinTabPage(VevoLAB);
            MeasurementTabPage.SearchProperties[WinTabPage.PropertyNames.Name] = " Measurement ";

            WinCheckBox ShowMeasurementsCheckBox = new WinCheckBox(VevoLAB);
            ShowMeasurementsCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Show Measurements";

            WinButton SaveAsButtonInMeasurement = new WinButton(VevoLAB);
            SaveAsButtonInMeasurement.SearchProperties[WinButton.PropertyNames.Name] = "Save As";

            WinTabPage AnnotationTabPage = new WinTabPage(VevoLAB);
            AnnotationTabPage.SearchProperties[WinTabPage.PropertyNames.Name] = "   Annotation   ";

            WinCheckBox ShowAnnotationsCheckBox = new WinCheckBox(VevoLAB);
            ShowAnnotationsCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Show Annotations";

            WinButton SaveAsButtonInAnnotation = new WinButton(VevoLAB);
            SaveAsButtonInAnnotation.SearchProperties[WinButton.PropertyNames.Name] = "Save As";

            WinEdit NewMeasurementPackageEditBox = new WinEdit(NewMeasurementPackagWindow);

            WinButton OKButtonInNewMeasurementPackageWindow = new WinButton(NewMeasurementPackagWindow);
            OKButtonInNewMeasurementPackageWindow.SearchProperties[WinButton.PropertyNames.Name] = "OK";

            WinButton YesButtonInPackageExistsWindow = new WinButton(PackageExistsWindow);
            YesButtonInPackageExistsWindow.SearchProperties[WinButton.PropertyNames.Name] = "Yes";

            WinButton OKButtonInPreferences = new WinButton(VevoLAB);
            OKButtonInPreferences.SearchProperties[WinButton.PropertyNames.Name] = "OK";

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

            WinCheckBox HideMeasurementsAndAnnotationsCheckBox = new WinCheckBox(VevoLAB);
            HideMeasurementsAndAnnotationsCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Hide measurements and annotations";

            WinRadioButton QualityRadioButton = new WinRadioButton(VevoLAB);
            QualityRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = TestContext.DataRow["Quality"].ToString();

            WinCheckBox ExportRegionsCheckBox = new WinCheckBox(VevoLAB);
            ExportRegionsCheckBox.SearchProperties[WinCheckBox.PropertyNames.Name] = "Export regions";

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
                if (strStudy != "" && SeriesName.Exists == false)
                {
                    Mouse.DoubleClick(StudyName, StudyName.GetClickablePoint());
                }

                //// Double-Click to expand the Series in the Study Browser
                if (strSeries != "" && ImageName.Exists == false)
                {
                    Mouse.DoubleClick(SeriesName, SeriesName.GetClickablePoint());
                }

                //// Double-Click the Image to load
                Mouse.DoubleClick(ImageName, ImageName.GetClickablePoint());

                //// Pause palyback and move to the first frame
                //// For Frames and 3D image (3D view), we do not need to stop playback
                if (PauseButton.Exists)
                {
                    ////// Pause playback
                    Mouse.Click(PauseButton, PauseButton.GetClickablePoint());

                    ////// Move to the first frame
                    Mouse.Click(FirstFrameButton, FirstFrameButton.GetClickablePoint());
                }

                // Switch to 'Image Processing' panel to set parameters
                // '3D Settings' panel has the same name ("PostProc") as 'Image Processing' panel
                if (ImageProcessingButton.Exists)
                {
                    Mouse.Click(ImageProcessingButton, ImageProcessingButton.GetClickablePoint());

                    //// Set 'Display Layout' combo box
                    if (TestContext.DataRow["Display Layout"].ToString() != "NA")
                    {
                        DisplayLayoutComboBox.SelectedItem = TestContext.DataRow["Display Layout"].ToString();
                    }

                    //// Clear 'Invert' check box
                    if (InvertCheckBox.Exists)
                    {
                        InvertCheckBox.Checked = false;
                    }

                    //// Set 'RF Overlay' check box
                    if (TestContext.DataRow["RF Overlay"].ToString() != "NA")
                    {
                        if (TestContext.DataRow["RF Overlay"].ToString() == "Enable")
                        {
                            RFOverlayCheckBox.Checked = true;
                        }
                        else if (TestContext.DataRow["RF Overlay"].ToString() == "Disable")
                        {
                            RFOverlayCheckBox.Checked = false;
                        }
                    }
                }

                // Switch to 'Physio' panel to set parameters
                if (PhysioButton.Exists)
                {
                    Mouse.Click(PhysioButton, PhysioButton.GetClickablePoint());

                    //// Set 'View Physiology' check box
                    if (TestContext.DataRow["View Physiology"].ToString() == "Enable")
                    {
                        ViewPhysiologyCheckBox.Checked = true;

                        ////// Select 'ECG' check box
                        ECGCheckBox.Checked = true;

                        ////// Select 'Respiration' check box
                        RespirationCheckBox.Checked = true;

                        ////// Clear 'Invert' check box
                        InvertRespirationCheckBox.Checked = false;

                        ////// Select 'BP' check box
                        BPCheckBox.Checked = true;

                        ////// Select 'BP Derivative' check box
                        BPDerivativeCheckBox.Checked = true;

                        ////// Select 'Temperature' check box
                        TemperatureCheckBox.Checked = true;
                    }
                    else if (TestContext.DataRow["View Physiology"].ToString() == "Disable")
                    {
                        ViewPhysiologyCheckBox.Checked = false;
                    }
                }

                // Switch to 'Preferences' window to set parameters
                Mouse.Click(PreferencesButton, PreferencesButton.GetClickablePoint());

                //// Switch to 'Preferences' window 'General' tab to set parameters
                Mouse.Click(GeneralTabPage, GeneralTabPage.GetClickablePoint());

                ////// Set 'Show Date/Time on Image Header' check box
                if (TestContext.DataRow["Show Date/Time on Image Header"].ToString() == "Enable")
                {
                    ShowDateTimeonImageHeaderCheckBox.Checked = true;
                }
                else if (TestContext.DataRow["Show Date/Time on Image Header"].ToString() == "Disable")
                {
                    ShowDateTimeonImageHeaderCheckBox.Checked = false;
                }

                //// Switch to 'Preferences' window 'Measurement' tab to set parameters
                Mouse.Click(MeasurementTabPage, MeasurementTabPage.GetClickablePoint());

                ////// Set 'Show Measurements' check box
                if (TestContext.DataRow["Show Measurements"].ToString() == "Enable")
                {
                    ShowMeasurementsCheckBox.Checked = true;
                }
                else if (TestContext.DataRow["Show Measurements"].ToString() == "Disable")
                {
                    ShowMeasurementsCheckBox.Checked = false;
                }

                ////// Click 'Save As' button which will open 'New Measurement Package' pop-up window
                Mouse.Click(SaveAsButtonInMeasurement, SaveAsButtonInMeasurement.GetClickablePoint());

                ////// Type 'Export Measurement Package' in edit box of 'New Measurement Package' pop-up window
                NewMeasurementPackageEditBox.Text = "Export Measurement Package";

                ////// Click 'OK' button in 'New Measurement Package' pop-up window
                Mouse.Click(OKButtonInNewMeasurementPackageWindow, OKButtonInNewMeasurementPackageWindow.GetClickablePoint());

                ////// Click 'Yes' button in 'Package Exists' pop-up window
                if (PackageExistsWindow.Exists)
                {
                    Mouse.Click(YesButtonInPackageExistsWindow, YesButtonInPackageExistsWindow.GetClickablePoint());
                }

                //// Switch to 'Preferences' window 'Annotation' tab to set parameters
                Mouse.Click(AnnotationTabPage, AnnotationTabPage.GetClickablePoint());

                ////// Set 'Show Annotations' check box
                if (TestContext.DataRow["Show Annotations"].ToString() == "Enable")
                {
                    ShowAnnotationsCheckBox.Checked = true;
                }
                else if (TestContext.DataRow["Show Annotations"].ToString() == "Disable")
                {
                    ShowAnnotationsCheckBox.Checked = false;
                }

                ////// Click 'Save As' button which will open 'New Measurement Package' pop-up window
                Mouse.Click(SaveAsButtonInAnnotation, SaveAsButtonInAnnotation.GetClickablePoint());

                ////// Type 'Export Measurement Package' in text box of 'New Measurement Package' pop-up window
                NewMeasurementPackageEditBox.Text = "Export Measurement Package";

                ////// Click 'OK' button in 'New Measurement Package' pop-up window
                Mouse.Click(OKButtonInNewMeasurementPackageWindow, OKButtonInNewMeasurementPackageWindow.GetClickablePoint());

                ////// Click 'Yes' button in 'Package Exists' pop-up window
                Mouse.Click(YesButtonInPackageExistsWindow, YesButtonInPackageExistsWindow.GetClickablePoint());

                ////// Click 'OK' button to close 'Preferences' window
                Mouse.Click(OKButtonInPreferences, OKButtonInPreferences.GetClickablePoint());
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

            //// Click 'Desktop' -> 'Export' tree item as the export destination
            Mouse.Click(ExportTreeItemInExportImageWindow, ExportTreeItemInExportImageWindow.GetClickablePoint());

            //// If it is a new 'Export Type'
            if (strExportType != "")
            {
                //// Click one of the 'Export Type' radio button
                Mouse.Click(ExportTypeRadioButton, ExportTypeRadioButton.GetClickablePoint());
            }

            //// Give a name in 'Save As' text box
            //// Export file name should not contain any space due to a requirement for "TIFF for 3D Volume Slices" format
            int TestResultRow = TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow);
            int ExcelRow = TestResultRow + 2;
            SaveAsInExportImageWindow.Text = string.Format("ExcelRow_{0}", ExcelRow);

            //// Select a 'File Type' from the combo box            
            FileTypeComboBox.SelectedItem = strFileType;

            //// Set 'Hide measurements and annotations' check box
            if (HideMeasurementsAndAnnotationsCheckBox.Exists)
            {
                if (TestContext.DataRow["Hide measurements and annotations"].ToString() == "Enable")
                {
                    HideMeasurementsAndAnnotationsCheckBox.Checked = true;
                }
                else if (TestContext.DataRow["Hide measurements and annotations"].ToString() == "Disable")
                {
                    HideMeasurementsAndAnnotationsCheckBox.Checked = false;
                }
            }

            //// Click one of the 'Quality' radio button
            if (QualityRadioButton.Exists)
            {
                //Mouse.Click(QualityRadioButton, QualityRadioButton.GetClickablePoint());
                QualityRadioButton.Selected = true;
            }

            //// Set 'Export regions' check box
            if (ExportRegionsCheckBox.Exists)
            {
                if (TestContext.DataRow["Export regions"].ToString() == "Enable")
                {
                    ExportRegionsCheckBox.Checked = true;
                }
                else if (TestContext.DataRow["Export regions"].ToString() == "Disable")
                {
                    ExportRegionsCheckBox.Checked = false;
                }
            }

            // Click 'OK' button in 'Export Image' window
            Mouse.Click(OKButtonInExportImageWindow, OKButtonInExportImageWindow.GetClickablePoint());

            // Click 'OK' button in 'Image Export Report' pop-up window
            Mouse.Click(OKButtonInImageExportReportWindow, OKButtonInImageExportReportWindow.GetClickablePoint());
        }

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
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
