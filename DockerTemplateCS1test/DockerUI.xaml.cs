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
                btn_drawSquad.Click += (s, e) => {
                    Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(@"E:\NextLeaf\test1.xlsx");
                    Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];
                    Microsoft.Office.Interop.Excel.Worksheet sheetoutput = workbook.Sheets[2];
                    int numRowsInput = sheet.UsedRange.Rows.Count;
                    int numRowsOutput = sheetoutput.UsedRange.Rows.Count;
                    int numColumns = 2;     // according to your sample

                    List<string> recordsInput = new List<string>();
                    List<string> recordsOutput = new List<string>();

                    Excel.Range cell;
                    int countInput = 0;
                    int countOutput = 0;

                    for (int rowIndex = 2; rowIndex <= numRowsInput; rowIndex++)
                    {
                        cell = (Excel.Range)sheet.Cells[rowIndex, 2];
                        if (Convert.ToString(cell.Value) != null)
                        {
                            recordsInput.Add(Convert.ToString(cell.Value));

                        }
                        else
                        {
                            recordsInput.Add("");

                        }
                    }
                    for (int rowIndexOutput = 2; rowIndexOutput <= numRowsOutput; rowIndexOutput++)
                    {
                        cell = (Excel.Range)sheetoutput.Cells[rowIndexOutput, 2];
                        if (Convert.ToString(cell.Value) != null)
                        {
                            recordsOutput.Add(Convert.ToString(cell.Value));
                        }
                        else
                        {
                            recordsOutput.Add("");
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
                                if (recordsInput[countInput] != "")
                                {
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, recordsInput[countInput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);

                                }
                                countInput++;
                                y -= 0.305;
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
                                if (recordsOutput[countOutput] != "")
                                {
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, recordsOutput[countOutput], (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 6, 0, 0, 0, (corel.cdrAlignment)0);
                                }
                                countOutput++;
                                y -= 0.395;
                            }
                            x += 1.64;
                        }

                    }

                };
                //popluate machpro zone
                btn_MachProZone.Click += (s, e) => {
                    int countLabel = 0;
                    //23 xls / 10 = 2.3 round up = 3 - 1 = 2 
                    string[] filePaths = Directory.GetFiles(@"F:\pointlist\machprozone");
                    //Pages pages = this.corelApp.ActiveDocument.Pages;
                    //Layers allLayers = pages[3].Layers;
                    ////copy and paste to other pages
                    //allLayers.Bottom.Shapes.All().Copy();
                    //int numpg = (int)(Math.Ceiling(filePaths.Length / 10.0)-1); 
                    //this.corelApp.ActiveDocument.InsertPages(numpg, false, 3);
                    //for (int j = 1; j < numpg+1; j++)
                    //{
                    //    pages[3+j].ActiveLayer.Paste();

                    //}

                    Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                    //for (int i = 0; i < filePaths.Length; i++)
                    for (int i = 0; i < 10; i++)
                    {
                        Microsoft.Office.Interop.Excel.Workbook workbook = xl.Workbooks.Open(@filePaths[i]);
                        Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.Sheets[1];
                        int numRowsInput = sheet.UsedRange.Rows.Count;
                        double xI, xO, yI = 8.46, yO = 8.4, xSS, ySS = 8.46;


                        Excel.Range cellPanelNumber;
                        cellPanelNumber = (Excel.Range)sheet.Cells[1, 2];

                        Excel.Range cellInOut;
                        Excel.Range cellLabel;
                        String strInOut;
                        xI = 2.48;
                        for (int rowIndex = 5; rowIndex <= numRowsInput; rowIndex++)
                        {
                            cellInOut = (Excel.Range)sheet.Cells[rowIndex, 1];
                            cellLabel = (Excel.Range)sheet.Cells[rowIndex, 2];
                            strInOut = cellInOut.Value;
                            if (strInOut != null && cellLabel.Value != null)
                            {
                                if (strInOut.ToLower().Contains('i'))
                                {
                                    if (countLabel == 5)
                                    {
                                        xI = 6.265;
                                    }
                                
                                    this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xI, yI, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xI, yI);
                                }
                                xI += 0.1624;
                                //else
                                //{
                                //    if (strInOut.ToLower().Contains('s'))
                                //    {
                                //        if (countLabel < 5)
                                //        {
                                //            xSS = 2.48 + 0.1624 * 7;
                                //        }
                                //        else
                                //        {
                                //            xSS = 6.265 + 0.1624 * 7;
                                //        }
                                    
                                //        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xSS, ySS, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xSS, ySS);
                                //    }
                                //    else
                                //    {
                                //        if (countLabel < 5)
                                //        {
                                //            xO = 0.98;
                                //        }
                                //        else
                                //        {
                                //            xO = 4.78;
                                //        }
                                    
                                //        this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(xO, yO, cellLabel.Value, (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, xO, yO);
                                //        xO += 0.1624;

                                //    }
                                //}

                            }
                        }
                                yI -= 1.65;

                                if (countLabel == 5)
                                {
                                    yI = 8.46;
                                    yO = 8.4;
                                    ySS = 8.46;
                                }
                                countLabel++;


                    }

                    xl.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);

                    ////--------------output------------------
                    //y = 8.46;
                    //x = 0.98;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "herree", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hellottttttttttt", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////bottom label output
                    //y = 8.46;
                    //x = 0.98;
                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "helloOUTCOl", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////right label output
                    //x = 4.78;
                    //y = 8.46;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "helloOUTCOl", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);


                    //--------------panel
                    //y = 9;
                    //x = 2.34;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "50102", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "50102", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    //y = 9;
                    //x += 3.78;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "50102", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "501024", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "501024", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "501024", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 8, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////--------------input
                    //y = 8.46;
                    //x = 2.48;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "herree", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hellottttttttttt", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
                    //x += 0.1624;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "hello", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////bottom label input
                    //y = 8.46;
                    //x = 2.48;
                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "helloINOl", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////right label input
                    //x = 6.265;
                    //y = 8.46;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "helloINOl", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);

                    ////right down label input
                    //x = 6.265;
                    //y = 8.46;
                    //y -= 1.65;
                    //this.corelApp.ActiveDocument.ActiveLayer.CreateArtisticText(x, y, "helloINOl", (corel.cdrTextLanguage)1033, 0, "Swis721 Cn BT", 7, 0, 0, 0, (corel.cdrAlignment)1).RotateEx(90.0, x, y);
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
