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

        public class Item
        {
            public string PanelName { get; set; }
            public List<string> recordsInputList { get; set; }
            public List<string> recordsOutputList { get; set; }
        }


        public DockerUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);


                void inputsIntoColumns(double x, double y, ref int countInput, double numInputRecords,  List<Item> itemsList, int j)
                {
                    //into 12 fields in input column
                    for (int i = 0; i < 12; i++)
                    {
                        if (countInput < numInputRecords)
                        {
                            if (itemsList[j].recordsInputList[countInput] != "")
                            {
                                this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, itemsList[j].recordsInputList[countInput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);

                            }
                            countInput++;
                            y -= 0.305;

                        }
                    }

                }                
                void outputIntoColumns(double x, double y,  ref int countOutput, double numOutputRecords, List<Item> itemsList, int j)
                {
                    //into 8 fields in output column
                    for (int k = 0; k < 8; k++)
                    {
                        if (countOutput < numOutputRecords)
                        {
                            if (itemsList[j].recordsOutputList[countOutput] != "")
                            {
                                this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, itemsList[j].recordsOutputList[countOutput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);
                            }
                            countOutput++;
                            y -= 0.395;
                        }
                    }

                }
                void numberInputsAndOutputs(double numX, double numY, int numPanelsPlaced)
                {
                    double numYCoor = numY;
                    double numXCoor = numX;
                    int num;
                    for (int xt = 1; xt < 13; xt++)
                    {
                        num = xt + 12 * (numPanelsPlaced - 1);
                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(numXCoor, numYCoor, "IN" + num, (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 8, corel.cdrTriState.cdrTrue, 0, 0, (corel.cdrAlignment)0);
                        numYCoor -= 0.304; 
                    }

                    //OUTPUT COLUMN
                    numYCoor = numY;
                    numXCoor = numX + 2.052;
                    for (int xt = 1; xt < 9; xt++)
                    {
                        num = xt + 8 * (numPanelsPlaced - 1);
                        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(numXCoor, numYCoor, "OUT" + num, (corel.cdrTextLanguage)1033, 0, "Swis721 BT", 8, corel.cdrTriState.cdrTrue, 0, 0, (corel.cdrAlignment)2);
                        numYCoor -= 0.3944;
                    }

                }

                    //popluate machpro sys labels
                    btn_MachProSys.Click += (s, e) => {
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
                        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                        int activePage = 1;
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                        Layers allLayers = pages[activePage].Layers;
                        //copy and paste to other pages
                        allLayers.Bottom.Shapes.All().Copy();
                        int numpg = filePaths.Length ;

                        List<Item> itemsList = new List<Item>();
                        List<string> recordsInput;
                        List<string> recordsOutput;
                        int countInput = 0;
                        int countOutput = 0;
                        string panelNameValue;
                        for (int w = 0; w < filePaths.Length; w++)
                        {

                            Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(@filePaths[w]);
                            Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];

                            int numRowsInput = sheet.UsedRange.Rows.Count;


                            Excel.Range cellInOut;
                            Excel.Range cellLabel;
                            Excel.Range PanelName = (Excel.Range)sheet.Cells[1, 2];
                            panelNameValue = "P" + PanelName.Value;
                            string strInOut;
                            recordsInput = new List<string>();
                            recordsOutput = new List<string>();
                            
                            //popluate into the list
                            for (int rowIndex = 5; rowIndex <= numRowsInput; rowIndex++)
                            {
                                cellInOut = (Excel.Range)sheet.Cells[rowIndex, 1];
                                cellLabel = (Excel.Range)sheet.Cells[rowIndex, 2];
                                strInOut = cellInOut.Value;

                                if (strInOut != null)
                                {
                                    if (strInOut.ToLower().Contains('i'))
                                    {
                                        if (cellLabel.Value != null)
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
                            itemsList.Add(new Item { PanelName = panelNameValue, recordsInputList = recordsInput, recordsOutputList = recordsOutput });

                            xl.Quit();
                        }
                        int cursorToLabel = 0;
                        int numPanelNames = 0;
                        double xPanel = 0;
                        double yPanel = 0;
                        double xNumbering = 0;
                        double yNumbering = 0;
                        bool hasChangePanel;
                        int numPanelsPlaced;
                        double numInputRecords;
                        double numOutputRecords;
                        //add the panel text at the top of the label
                        for (int j = 0; j < itemsList.Count(); j++)
                        {

                            hasChangePanel = true;
                            numPanelsPlaced = 0;
                            numInputRecords = itemsList[j].recordsInputList.Count();
                            numOutputRecords = itemsList[j].recordsOutputList.Count();
                            if (numInputRecords > numOutputRecords)
                            {
                                numPanelNames = (int)(Math.Ceiling(numInputRecords / 12) );
                            }
                            else
                            {
                                numPanelNames = (int)(Math.Ceiling(numOutputRecords / 8));
                            }
                            xPanel = 1.055;
                            x = 0.50;
                            xNumbering = 0.525;
                            while (numPanelsPlaced < numPanelNames)
                            {
                                numPanelsPlaced++;
                                if (cursorToLabel < 3)
                                {
                                    yPanel = 9.25;
                                    y = 8.88;
                                    yNumbering = 8.987;
                                }
                                else
                                {
                                    yPanel = 5.25;
                                    y = 4.90;
                                    yNumbering = 5.0115;
                                }
                                if (hasChangePanel)
                                {
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xPanel, yPanel, itemsList[j].PanelName + " MACH-Pro-Sys 1", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, corel.cdrTriState.cdrTrue, 0, 0, (corel.cdrAlignment)0);
                                    hasChangePanel = false;
                                }
                                else
                                {
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xPanel, yPanel, itemsList[j].PanelName + " MACH-Pro-Point " + numPanelsPlaced, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, corel.cdrTriState.cdrTrue, 0, 0, (corel.cdrAlignment)0);
                                }

                                inputsIntoColumns(x, y, ref countInput, numInputRecords, itemsList, j);
                                x += 1.078;
                                outputIntoColumns(x, y,  ref countOutput,  numOutputRecords, itemsList, j);
                                x += 1.64;

                                numberInputsAndOutputs(xNumbering, yNumbering, numPanelsPlaced);
                                xNumbering += 2.715;
                                xPanel += 2.678;
                                if (cursorToLabel == 2)
                                {
                                    xPanel = 1.055;
                                    x = 0.50;
                                    xNumbering = 0.525;
                                }
                                cursorToLabel++;
                            }

                            //add the inputs and output to the labels
                            
                            countInput = 0;
                            countOutput = 0;
                            //change page when the inputs / output are full
                            if (cursorToLabel == 6)
                            {
                                cursorToLabel = 0;
                                this.corelApp.ActiveDocument.InsertPages(1, false, activePage);
                                pages[activePage+1].ActiveLayer.Paste();
                                pages[++activePage].Activate();
                            }
                        }
                        
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
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
                        int activePage = 1;
                        Pages pages = this.corelApp.ActiveDocument.Pages;
                        Layers allLayers = pages[activePage].Layers;
                        //copy and paste to other pages
                        allLayers.Bottom.Shapes.All().Copy();
                        int numpg = (int)(Math.Ceiling(filePaths.Length / 10.0) - 1);
                        this.corelApp.ActiveDocument.InsertPages(numpg, false, activePage);
                        for (int j = 1; j < numpg + 1; j++)
                        {
                            pages[activePage + j].ActiveLayer.Paste();

                        }
                        pages[activePage].Activate();
                        Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                        
                        double xI = 2.17, xO, xp = 2.34, yp = 9, yI = 8.335, yO = 8.415, xSS, ySS;
                        int numMach = 0;
                        for (int i = 0; i < filePaths.Length; i++)
                        {
                            if (numMach == 10)
                            {
                                pages[++activePage].Activate();
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
                //export pages to png files
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

                //export pages to 480x272 png files
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
                // inputs/ outputs points label 
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
                        string[] filePaths = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories).Where(name => !name.Contains("~$")).ToArray();

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
                        int activePg = 1;
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
