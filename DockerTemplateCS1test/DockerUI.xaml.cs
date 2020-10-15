using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using corel = Corel.Interop.VGCore;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Corel.Interop.VGCore;


namespace DockerTemplateCS1test
{
    public partial class DockerUI : UserControl
    {
        private corel.Application corelApp;
        private Styles.StylesController stylesController;
        private corel.Shape ll;
        private corel.Shape kk;
        private double x;
        private double y;
        public DockerUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
                //popluate machpro sys
                btn_MachProSys.Click += (s, e) => {
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    folderBrowser.InitialDirectory = "c:\\";
                    folderBrowser.Filter = "Database files (*.xlsx)| *.xlsx";
                    folderBrowser.FilterIndex = 0;
                    folderBrowser.RestoreDirectory = true;
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = folderBrowser.FileName;
                        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(folderPath);
                        Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];

                        int numRowsInput = sheet.UsedRange.Rows.Count;
   
                        List<string> recordsInput = new List<string>();
                        List<string> recordsOutput = new List<string>();

                        Excel.Range cell;
                        Excel.Range cellInOut;
                        Excel.Range cellLabel;
                        string strInOut;
                        int countInput = 0;
                        int countOutput = 0;
                        //popluate into the list
                        for (int rowIndex = 5; rowIndex <= numRowsInput; rowIndex++)
                        {
                            cellInOut = (Excel.Range)sheet.Cells[rowIndex, 1];
                            cellLabel = (Excel.Range)sheet.Cells[rowIndex, 2];
                            strInOut = cellInOut.Value;

                            if (strInOut != null )
                            {
                                if (strInOut.ToLower().Contains('i'))
                                {
                                    if(cellLabel.Value != null)
                                    {
                                        recordsInput.Add(Convert.ToString(cellLabel.Value));

                                    }
                                    else
                                    {
                                        recordsInput.Add("");
                                    }
                                }
                                else
                                {
                                    if (cellLabel.Value != null)
                                    {
                                        recordsOutput.Add(Convert.ToString(cellLabel.Value));

                                    }
                                    else
                                    {
                                        recordsOutput.Add("");
                                    }

                                }
                            }

                        }

                        xl.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);

                        for (int l = 0; l < 2; l++)
                        {
                            x = 0.50;
                            for (int j = 0; j < 3; j++)
                            {
                                if (l == 0)
                                {
                                    y = 8.88;
                                }
                                else
                                {
                                    y = 4.90;
                                }
                                for (int i = 0; i < 12; i++)
                                {
                                    if (countInput < recordsInput.Count())
                                    {
                                        if (recordsInput[countInput] != "")
                                        {
                                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, recordsInput[countInput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);

                                        }
                                        countInput++;
                                        y -= 0.305;

                                    }
                                }

                                x += 1.078;
                                if (l == 0)
                                {
                                    y = 8.88;
                                }
                                else
                                {
                                    y = 4.90;
                                }
                                for (int k = 0; k < 8; k++)
                                {
                                    if (countOutput < recordsOutput.Count())
                                    {
                                        if (recordsOutput[countOutput] != "")
                                        {
                                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, recordsOutput[countOutput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);
                                        }
                                        countOutput++;
                                        y -= 0.395;
                                    }
                                }
                                x += 1.64;
                            }

                        }
                    }

                };

                //popluate machpro zone
                btn_MachProZone.Click += (s, e) => {
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    // Set validate names and check file exists to false otherwise windows will
                    // not let you select "Folder Selection."
                    folderBrowser.ValidateNames = false;
                    folderBrowser.CheckFileExists = false;
                    folderBrowser.CheckPathExists = true;
                    // Always default to Folder Selection.
                    folderBrowser.FileName = "Folder Selection.";
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                        int countLabel = 0;
                        string[] filePaths = Directory.GetFiles(folderPath).Where(name => !name.Contains("~$")).ToArray();
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                        Layers allLayers = pages[3].Layers;
                        //copy and paste to other pages
                        allLayers.Bottom.Shapes.All().Copy();
                        int numpg = (int)(Math.Ceiling(filePaths.Length / 10.0) - 1);
                        this.corelApp.ActiveDocument.InsertPages(numpg, false, 3);
                        for (int j = 1; j < numpg + 1; j++)
                        {
                            pages[3 + j].ActiveLayer.Paste();

                        }
                        pages[3].Activate();
                        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                        
                        double xI = 2.17, xO, xp = 2.34, yp = 9, yI = 8.335, yO = 8.415, xSS, ySS;
                        int numMach = 0, activePg = 3;
                        for (int i = 0; i < filePaths.Length; i++)
                        {
                            if (numMach == 10)
                            {
                                pages[++activePg].Activate();
                                numMach = 0;
                                yp = 9; yI = 8.335; yO = 8.415; ySS = 8.46;
                                countLabel = 0;
                            }

                            numMach++;
                            Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(@filePaths[i]);
                            Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];
                            int numRowsInput = sheet.UsedRange.Rows.Count;


                            Excel.Range cellPanelNumber;
                            cellPanelNumber = (Excel.Range)sheet.Cells[1, 2];

                            Excel.Range cellInOut;
                            Excel.Range cellLabel;
                            String strInOut;


                            if (countLabel >= 5)
                            {
                                xI = 6.265;
                                xp = 6.12;
                                xO = 4.76;
                                xSS = 6.265 + 0.1624 * 7;
                            }
                            else
                            {
                                xI = 2.48;
                                xp = 2.34;
                                xO = 0.98;
                                xSS = 2.48 + 0.1624 * 7;
                            }

                            if (countLabel == 5)
                            {
                                yI = 8.335;
                                yp = 9;
                                yO = 8.415;
                                ySS = 8.335;
                            }
                            bool empty = false;
                            //put panel number to coreldraw
                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xp, yp, cellPanelNumber.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xp, yp);
                            for (int rowIndex = 5; rowIndex <= numRowsInput; rowIndex++)
                            {
                                cellInOut = (Excel.Range)sheet.Cells[rowIndex, 1];
                                cellLabel = (Excel.Range)sheet.Cells[rowIndex, 2];
                                strInOut = cellInOut.Value;

                                if (strInOut != null && cellLabel.Value != null)
                                {
                                    if (strInOut.ToLower().Contains('i'))
                                    {
                                        if (empty)
                                        {
                                            xI += 0.1624;
                                        }
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xI, yI, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 5, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xI, yI);
                                        xI += 0.1624;
                                        empty = false;
                                    }
                                    else
                                    {
                                        if (strInOut.ToLower().Contains('s'))
                                        {
                                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xSS, yI, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 5, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xSS, yI);
                                        }
                                        else
                                        {
                                            if (empty)
                                            {
                                                xO += 0.1624;
                                            }
                                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xO, yO, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 5, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xO, yO);
                                            xO += 0.1624;
                                            empty = false;
                                        }
                                    }
                                }
                                else
                                {
                                    empty = true;
                                }
                            }
                            yI -= 1.65;
                            yp -= 1.65;
                            yO -= 1.65;
                            countLabel++;

                            xl.Quit();
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                    }
                };

                btn_ExportAllPgPNG.Click += (s, e) =>
                {
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    // Set validate names and check file exists to false otherwise windows will
                    // not let you select "Folder Selection."
                    folderBrowser.ValidateNames = false;
                    folderBrowser.CheckFileExists = false;
                    folderBrowser.CheckPathExists = true;
                    // Always default to Folder Selection.
                    folderBrowser.FileName = "Folder Selection.";
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                        string name;
                        List<string> listnamespg = new List<string>();
                        StructExportOptions opt = new StructExportOptions();
                        for (int j = 1; j < pages.Count + 1; j++)
                        {
                            name = pages[j].Name.ToLower();

                            name = name.Replace(' ', '_');
                            listnamespg.Add(name);
                        }

                        opt.AntiAliasingType = cdrAntiAliasingType.cdrNormalAntiAliasing;
                        opt.Transparent = true;
                        opt.ImageType = cdrImageType.cdrRGBColorImage;
                        opt.ResolutionX = 300;
                        opt.ResolutionY = 300;
                        opt.MaintainAspect = true;
                        opt.SizeX = 1200;
                        opt.SizeY = 800;
                        int activepage;
                        for (int k = 0; k < listnamespg.Count; k++)
                        {
                            activepage = k + 1;
                            pages[activepage].Activate();
                            this.corelApp.ActiveDocument.Export(folderPath + "\\" + listnamespg[k] + ".png", cdrFilter.cdrPNG, cdrExportRange.cdrCurrentPage, opt);
                        }
                    }
                };


                btn_ExportAllPgPNG_Proview.Click += (s, e) =>
                {
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    // Set validate names and check file exists to false otherwise windows will
                    // not let you select "Folder Selection."
                    folderBrowser.ValidateNames = false;
                    folderBrowser.CheckFileExists = false;
                    folderBrowser.CheckPathExists = true;
                    // Always default to Folder Selection.
                    folderBrowser.FileName = "Folder Selection.";
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                        string name;
                        List<string> listnamespg = new List<string>();
                        StructExportOptions opt = new StructExportOptions();
                        for (int j = 1; j < pages.Count + 1; j++)
                        {
                            name = pages[j].Name.ToLower();

                            name = name.Replace(' ', '_');
                            listnamespg.Add(name);
                        }

                        opt.AntiAliasingType = cdrAntiAliasingType.cdrNormalAntiAliasing;
                        opt.Transparent = true;
                        opt.ImageType = cdrImageType.cdrRGBColorImage;
                        opt.ResolutionX = 300;
                        opt.ResolutionY = 300;
                        opt.MaintainAspect = true;
                        opt.SizeX = 480;
                        opt.SizeY = 272;
                        int activepage;
                        for (int k = 0; k < listnamespg.Count; k++)
                        {
                            activepage = k + 1;
                            pages[activepage].Activate();
                            this.corelApp.ActiveDocument.Export(folderPath + "\\" + listnamespg[k] + ".png", cdrFilter.cdrPNG, cdrExportRange.cdrCurrentPage, opt);
                        }
                    }
                };

                btn_InputOutputs.Click += (s, e) =>
                {
                    Dictionary<string, string> records = new Dictionary<string, string>();
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    // Set validate names and check file exists to false otherwise windows will
                    // not let you select "Folder Selection."
                    folderBrowser.ValidateNames = false;
                    folderBrowser.CheckFileExists = false;
                    folderBrowser.CheckPathExists = true;
                    // Always default to Folder Selection.
                    folderBrowser.FileName = "Folder Selection.";
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                        string[] filePaths = Directory.GetFiles(folderPath).Where(name => !name.Contains("~$")).ToArray();

                        for (int i = 0; i < filePaths.Length; i++)
                        {
                            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(@filePaths[i]);
                            Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];
                            int numRowsInput = sheet.UsedRange.Rows.Count;

                            Excel.Range cellInOut;
                            Excel.Range cellLabel;
                            String strInOut;

                            //put panel number to coreldraw
                            string strLabel;
                            for (int rowIndex = 5; rowIndex <= numRowsInput; rowIndex++)
                            {
                                cellInOut = (Excel.Range)sheet.Cells[rowIndex, 1];
                                cellLabel = (Excel.Range)sheet.Cells[rowIndex, 2];
                                strInOut = cellInOut.Value;

                                if (Convert.ToString(cellLabel.Value) != null)
                                {
                                    //add a space between the number and letter
                                    strLabel = "[P" + Convert.ToString(cellInOut.Value) + "]";
                                    records.Add(strLabel, Convert.ToString(cellLabel.Value));

                                }

                            }

                            xl.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                        }

                        int countInput = 0;
                        int recCount = records.Count();
                        int activePg = 2;
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                    
                        Layers allLayers = pages[activePg].Layers;
                        //copy and paste to other pages
                        allLayers.Bottom.Shapes.All().Copy();
                        //number of columns
                        int numCol = (int)(Math.Ceiling(recCount / 16.0));
                        int numPg = (int)(Math.Ceiling(numCol / 3.0) - 1);
                        this.corelApp.ActiveDocument.InsertPages(numPg, false, activePg);
                        for (int j = 1; j < numPg + 1; j++)
                        {
                            pages[activePg + j].ActiveLayer.Paste();

                        }

                        pages[activePg].Activate();
                        //16 rows, 3 columns
                        double x = 1.64, y= 9.83;
                        for(int j = 0; j < numCol; j++)
                        {
                            if(j % 3 == 0)
                            {
                                pages[activePg++].Activate();
                                x = 1.64; y = 9.83;
                            }
                            for (int i = 0; i < 16; i++)
                            {
                                if (countInput < recCount)
                                {
                                    var item = records.ElementAt(countInput);
                                    string itemKey = item.Key;
                                    string itemValue = item.Value;
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, itemValue, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)3);
                                    y -= 0.1;
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, itemKey, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)3);
                                    y -= 0.498;

                                }
                                countInput++;
                            }
                            x += 2.07;
                            y = 9.83;
                        }
                    }
                };
                btn_TAGS.Click += (s, e) =>
                {
                    var buildingValueList = new List<string>();
                    var panelValueValueList = new List<string>();
                    var descValueValueList = new List<string>();
                    var networkValueList = new List<string>();
                    var bacnetIdValueList = new List<string>();
                    var pointValueList = new List<string>();
                    var tagNameValueList = new List<string>();
                    var wireValueList = new List<string>();
                    System.Windows.Forms.OpenFileDialog folderBrowser = new System.Windows.Forms.OpenFileDialog();

                    folderBrowser.InitialDirectory = "c:\\";
                    folderBrowser.Filter = "Database files (*.xlsx)| *.xlsx";
                    folderBrowser.FilterIndex = 0;
                    folderBrowser.RestoreDirectory = true;
                    if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string folderPath = folderBrowser.FileName;
                        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(folderPath);
                        Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];
                        int numRowsInput = sheet.UsedRange.Rows.Count;

                        Excel.Range cellBuilding;
                        Excel.Range cellPanel;
                        Excel.Range cellDesc;
                        Excel.Range cellNetwork;
                        Excel.Range cellBacnetID;
                        Excel.Range cellPoints;
                        Excel.Range cellTagName;
                        Excel.Range cellWired;

                        //put panel number to coreldraw
                        string buildingValue ="none";
                        string panelValue ="none";
                        string descValue ="none";
                        string networkValue ="none";
                        string bacnetIdValue ="none";
                        string pointValue ="none";
                        string tagNameValue ="none";
                        string wireValue ="none";
                        for (int rowIndex = 2; rowIndex <= numRowsInput; rowIndex++)
                        {
                            cellBuilding = (Excel.Range)sheet.Cells[rowIndex, 1];
                            cellPanel = (Excel.Range)sheet.Cells[rowIndex, 2];
                            cellDesc = (Excel.Range)sheet.Cells[rowIndex, 3];
                            cellNetwork = (Excel.Range)sheet.Cells[rowIndex, 4];
                            cellBacnetID = (Excel.Range)sheet.Cells[rowIndex, 5];
                            cellPoints = (Excel.Range)sheet.Cells[rowIndex, 6];
                            cellTagName = (Excel.Range)sheet.Cells[rowIndex, 7];
                            cellWired = (Excel.Range)sheet.Cells[rowIndex, 8];



                            //add a space between the number and letter
                            buildingValueList.Add(Convert.ToString(cellBuilding.Value));
                            panelValueValueList.Add(Convert.ToString(cellPanel.Value));
                            descValueValueList.Add(Convert.ToString(cellDesc.Value));
                            networkValueList.Add(Convert.ToString(cellNetwork.Value));
                            bacnetIdValueList.Add(Convert.ToString(cellBacnetID.Value));
                            pointValueList.Add(Convert.ToString(cellPoints.Value));
                            tagNameValueList.Add(Convert.ToString(cellTagName.Value));
                            wireValueList.Add(Convert.ToString(cellWired.Value));

                        }

                        xl.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);

                        
                        
                        int recCount = buildingValueList.Count();
                        int activePg = 1;
                        Pages pages = this.corelApp.ActiveDocument.Pages;

                        Layers allLayers = pages[activePg].Layers;
                        //copy and paste to other pages
                        allLayers.Bottom.Shapes.All().Copy();
                        //number of columns
                        int numPg = (int)(Math.Ceiling(recCount / 12.0)-1);
                        
                        this.corelApp.ActiveDocument.InsertPages(numPg, false, activePg);
                        for (int j = 1; j < numPg + 1; j++)
                        {
                            pages[activePg + j].ActiveLayer.Paste();

                        }
                        int recordIndex = 0;
                        for (int p = 0; p < numPg+1; p++)
                        {
                            double y = 9.105;
                            pages[activePg++].Activate();
                            for (int k = 0; k < 3; k++)
                            {
                                double x = 1.10, xMid = 1.2, xBottom = 0.8;

                                for (int i = 0; i < 4; i++)
                                {
                                    switch (k)
                                    {
                                        case 0:
                                            y = 9.105;
                                            break;
                                        case 1:
                                            y = 5.655;
                                            break;
                                        case 2:
                                            y = 2.205;
                                            break;
                                        default:
                                            break;
                                    }
                                    if (recordIndex < buildingValueList.Count())
                                    {
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, buildingValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        y = y - 0.095;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, panelValueValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        y = y - 0.095;
                                        if (descValueValueList[recordIndex] != null)
                                        {
                                            this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, descValueValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        }
                                        //put the text at network and bacnet id white boxes
                                        y = y - 0.4;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xMid, y, networkValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 9, 0, 0, 0, (corel.cdrAlignment)1);
                                        y = y - 0.155;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xMid, y, bacnetIdValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 9, 0, 0, 0, (corel.cdrAlignment)1);

                                        //point, name, and wired white boxes
                                        y = y - 0.445;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xBottom, y, pointValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        y = y - 0.18;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xBottom, y, tagNameValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        y = y - 0.18;
                                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xBottom, y, wireValueList[recordIndex], (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 7, 0, 0, 0, (corel.cdrAlignment)1);
                                        
                                        recordIndex++;
                                        x = x + 2.02;
                                        xMid = xMid + 2.02;
                                        xBottom = xBottom + 2.02;

                                    }

                                }
                            }
                            
                        }


                    }


                };

            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }

        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
        }
    }
}
