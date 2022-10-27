using CommunityToolkit.HighPerformance;
using Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomisationHospitalData
{
    public partial class Form1 : Form
    {
        // Objects for the merged worksheet
        Excel.Application excelProgram;
        _Workbook workbookMerged;
        _Worksheet worksheetMerged;
        Range rangeMerged;

        // List of String arrays for the category library
        List<string[]> listLibrary = new List<string[]>();

        // File dialogues
        private OpenFileDialog openBibliotekPathDialog = new OpenFileDialog();
        private OpenFileDialog openACPathDialog = new OpenFileDialog();
        private OpenFileDialog openBCPathDialog = new OpenFileDialog();
        private OpenFileDialog openCBPBageriPathDialog = new OpenFileDialog();
        private OpenFileDialog openDagrofaPathDialog = new OpenFileDialog();
        private OpenFileDialog openEmmerysPathDialog = new OpenFileDialog();
        private OpenFileDialog openFrisksnitPathDialog = new OpenFileDialog();
        private OpenFileDialog openGrøntGrossistenPathDialog = new OpenFileDialog();
        private OpenFileDialog openHørkramPathDialog = new OpenFileDialog();

        private OpenFileDialog openPathDialog = new OpenFileDialog();

        string pathBibliotek = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\Kategoribibliotek 2.xlsx";

        // paths for companies supplying folders of excel sheets
        List<string> pathAC = new List<string>();
        List<string> pathDagrofa = new List<string>();
        List<string> pathFrisksnit = new List<string>();
        List<string> pathDeViKas = new List<string>();

        // paths for companies supplying individual excel sheets
        string pathBC = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\BC.xlsx";
        string pathCBPBageri = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\CBP bageri.xlsx";
        string pathEmmerys = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\Emmerys 01-04-2021..30-06-2021.xlsx";
        string pathGrøntGrossisten = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\Grønt Grossisten.xlsx";
        string pathHørkram = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\Hørkram.xlsx";

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //Start Excel and get Application object.
            excelProgram = new Excel.Application();
            excelProgram.Visible = true;

            //Get a new workbook.
            workbookMerged = (Excel._Workbook)(excelProgram.Workbooks.Add(Missing.Value));
            worksheetMerged = (Excel._Worksheet)workbookMerged.ActiveSheet;

            //Add table headers going cell by cell.
            worksheetMerged.Cells[1, 1] = "År";
            worksheetMerged.Cells[1, 2] = "Kvartal";
            worksheetMerged.Cells[1, 3] = "Hospital";
            worksheetMerged.Cells[1, 4] = "Råvarekategori";
            worksheetMerged.Cells[1, 5] = "Leverandør";
            worksheetMerged.Cells[1, 6] = "Råvare";
            worksheetMerged.Cells[1, 7] = "konv/øko";
            worksheetMerged.Cells[1, 8] = "Varianter/opr";
            worksheetMerged.Cells[1, 9] = "Pris pr enhed";
            worksheetMerged.Cells[1, 10] = "Pris i alt";
            worksheetMerged.Cells[1, 11] = "Kg";
            worksheetMerged.Cells[1, 12] = "Kilopris";
            worksheetMerged.Cells[1, 13] = "Oprindelse";
            worksheetMerged.Cells[1, 14] = "Kg CO2-eq pr kg";
            worksheetMerged.Cells[1, 15] = "Kg CO2-eq pr kg total";
            worksheetMerged.Cells[1, 16] = "Kg CO2-eq pr MJ";
            worksheetMerged.Cells[1, 17] = "Kg CO2-eq pr MJ total";
            worksheetMerged.Cells[1, 18] = "Kg CO2-eq pr g protein";
            worksheetMerged.Cells[1, 19] = "Kg CO2-eq pr g protein total";
            worksheetMerged.Cells[1, 20] = "Arealanvendelse m2";
            worksheetMerged.Cells[1, 21] = "Arealanvendelse m2 total";
            worksheetMerged.Cells[1, 22] = "CO2 tal";

            //Format A1:V1 as bold, vertical alignment = center.
            worksheetMerged.get_Range("A1", "V1").Font.Bold = true;
            worksheetMerged.get_Range("A1", "V1").Interior.Color = XlRgbColor.rgbSlateBlue;
            worksheetMerged.get_Range("A1", "V1").Font.Color = XlRgbColor.rgbWhite;
            worksheetMerged.get_Range("A1", "V1").VerticalAlignment =
            Excel.XlVAlign.xlVAlignCenter;
        }
        private void ButtonBibliotekPath_Click(object sender, EventArgs e)
        {
            this.openBibliotekPathDialog.Title = "Select Bibliotek File";
            if (openBibliotekPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathBibliotek = openBibliotekPathDialog.FileName;
                ButtonBibliotekPath.Text = openBibliotekPathDialog.FileName;
            }
        }
        private void CreateBibliotekButton_Click(object sender, System.EventArgs e)
        {
            Excel._Workbook workbookLibrary;
            Excel._Worksheet worksheetLibrary;
            Excel.Range rangeLibrary;

            workbookLibrary = excelProgram.Workbooks.Open(pathBibliotek);
            worksheetLibrary = workbookLibrary.Sheets[1];
            rangeLibrary = worksheetLibrary.UsedRange;

            int rowCountLibrary = rangeLibrary.Rows.Count;
            int colCountLibrary = rangeLibrary.Columns.Count;

            Object[,] arrayLibrary = rangeLibrary.get_Value();

            for (int row = 0; row < rowCountLibrary - 2; row++)
            {
                listLibrary.Add(new string[5]);
                for (int col = 0; col < colCountLibrary; col++)
                {
                    listLibrary[row].SetValue(arrayLibrary[row + 3, col + 1].ToString(), col);
                }
            }

            listLibrary = listLibrary.Distinct().ToList();

            workbookLibrary.Close(false);
            MRCO(workbookLibrary);
            MRCO(worksheetLibrary);
            MRCO(rangeLibrary);
        }

        // Code for AC files
        private void ButtonACPath_Click(object sender, EventArgs e)
        {
            this.openACPathDialog.Multiselect = true;
            this.openACPathDialog.Title = "Select AC files";

            if (openACPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathAC = openACPathDialog.FileNames.ToList();
                ButtonACPath.Text = openACPathDialog.FileName;
            }
        }

        // Code for BC files
        private void ButtonBCPath_Click(object sender, EventArgs e)
        {
            this.openBCPathDialog.Title = "Select BC File";
            if (openBCPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathBC = openBCPathDialog.FileName;
                ButtonBCPath.Text = openBCPathDialog.FileName;
            }
        }

        // Code for CBP Bageri files
        private void ButtonCBPBageriPath_Click(object sender, EventArgs e)
        {
            this.openCBPBageriPathDialog.Title = "Select CBP Bageri File";
            if (openCBPBageriPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathCBPBageri = openCBPBageriPathDialog.FileName;
                ButtonCBPBageriPath.Text = openCBPBageriPathDialog.FileName;
            }
        }
        private void CBPBageriButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookCBP;
            Excel._Worksheet worksheetCBP;
            Excel._Worksheet infosheetCBP;
            Excel.Range rangeCBP;

            try
            {
                workbookCBP = excelProgram.Workbooks.Open(pathCBPBageri);
                infosheetCBP = workbookCBP.Sheets[1];
                worksheetCBP = workbookCBP.Sheets[2];
                rangeCBP = worksheetCBP.UsedRange;

                int headerRows = 1;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountCBP = rangeCBP.Rows.Count - headerRows;
                int colCountCBP = rangeCBP.Columns.Count;

                // Imports the date from the CBP worksheet
                string stringDate = infosheetCBP.Cells[1, 1].Text;
                DateTime dateCBP = DateTime.Parse(stringDate.Split(' ')[3]);

                // Imports the cell data from the CBP sheet as an array of Objects
                Object[,] arrayCBP = rangeCBP.get_Value();

                // Creates a List of String arrays for every row in the CBP worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<String[]> listCBP = new List<String[]>();

                // For every row in the imported CBP Object array, copy its value to the corresponding String in the List of String arrays
                for (int row = 0; row < rowCountCBP; row++)
                {
                    listCBP.Add(new string[14]);
                    for (int col = 0; col < colCountCBP; col++)
                    {
                        try // "Try" because the cell's value can be Null
                        {
                            listCBP[row].SetValue(arrayCBP[row + 1 + headerRows, col + 1].ToString(), col);
                        }
                        catch (NullReferenceException) // "Catch" in case the cell's value is Null
                        {
                            listCBP[row].SetValue("", col);
                        }
                    }
                }

                // Deletion of irrelevant entries from the List of String arrays
                listCBP.RemoveRange(0, 2); // Header entries in row 1 and 2

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listCBP.Count));

                object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                // Sets the values in the CBP Object Array
                for (int row = 0; row < listCBP.Count; row++)
                {
                    arrayMerged[row + 1, 1] = dateCBP.Year; // År
                    arrayMerged[row + 1, 2] = dateCBP.Month / 3; // Kvartal
                    arrayMerged[row + 1, 3] = listCBP[row].GetValue(0).ToString().Split(new string[] { " ~ " }, StringSplitOptions.None).Last(); // Hospital
                    arrayMerged[row + 1, 4] = ""; // Råvarekategori
                    arrayMerged[row + 1, 5] = "CBP Bageri"; // Leverandør
                    arrayMerged[row + 1, 6] = ""; // Råvare
                    if (listCBP[row].GetValue(10) as String == "Økologi") // konv/øko
                    {
                        arrayMerged[row + 1, 7] = "Øko";
                    }
                    if (listCBP[row].GetValue(10) as String == "Ej Økologi")
                    {
                        arrayMerged[row + 1, 7] = "Konv";
                    }
                    arrayMerged[row + 1, 8] = listCBP[row].GetValue(1).ToString().Split(new string[] { " ~ " }, StringSplitOptions.None).Last(); // Varianter/opr
                    arrayMerged[row + 1, 9] = ""; // Pris pr enhed
                    arrayMerged[row + 1, 10] = listCBP[row].GetValue(3); // Pris i alt
                    arrayMerged[row + 1, 11] = listCBP[row].GetValue(2); // Kg
                    arrayMerged[row + 1, 12] = listCBP[row].GetValue(8); // Kilopris
                    arrayMerged[row + 1, 13] = ""; // Oprindelse
                }

                rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listCBP.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listCBP.Count)).Font.Size = 11;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();


                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects
                workbookCBP.Close(false);
                MRCO(workbookCBP);
                MRCO(worksheetCBP);
                MRCO(rangeCBP);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Code for Dagrofa files
        private void ButtonDagrofaPath_Click(object sender, EventArgs e)
        {
            this.openDagrofaPathDialog.Multiselect = true;
            this.openDagrofaPathDialog.Title = "Select Dagrofa files";

            if (openDagrofaPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathDagrofa = openDagrofaPathDialog.FileNames.ToList();
                ButtonDagrofaPath.Text = openDagrofaPathDialog.FileName;
            }
        }
        private void DagrofaButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookDagrofa;
            Excel._Worksheet worksheetDagrofa;
            Excel.Range rangeDagrofa;

            try
            {
                foreach (string fileDagrofa in pathDagrofa)
                {
                    workbookDagrofa = excelProgram.Workbooks.Open(fileDagrofa);
                    worksheetDagrofa = workbookDagrofa.Sheets[1];
                    rangeDagrofa = worksheetDagrofa.UsedRange;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    int rowCountDagrofa = rangeDagrofa.Rows.Count;
                    int colCountDagrofa = rangeDagrofa.Columns.Count;

                    int headerRows = 9;
                    int headerCols = 5;

                    object[,] arrayDagrofa = rangeDagrofa.get_Value();

                    string[] date = arrayDagrofa[1, 1].ToString().Split(' ');
                    string year = date[date.Length - 1];
                    string quarter = date[date.Length - 2].Replace("Q", "");

                    List<string[]> listDagrofa = new List<string[]>();
                    try
                    {
                        for (int row = headerRows; row < rowCountDagrofa - 1; row++)
                        {
                            for (int col = headerCols; col < colCountDagrofa; col += 4)
                            {
                                try
                                {
                                    float currentAmount = float.Parse(arrayDagrofa[row, col].ToString());
                                    if (currentAmount > 0)
                                    {
                                        listDagrofa.Add(new string[13]);

                                        listDagrofa[listDagrofa.Count - 1].SetValue(year, 0); // År
                                        if (listDagrofa[listDagrofa.Count - 1].GetValue(0) == null)
                                            listDagrofa[listDagrofa.Count - 1].SetValue(quarter, 1); // Kvartal
                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[7, col], 2); // Hospital

                                        string[] råvare = GetRåvare("Dagrofa", arrayDagrofa[row, 2].ToString());

                                        listDagrofa[listDagrofa.Count - 1].SetValue(råvare[0], 3); // Råvarekategori
                                        listDagrofa[listDagrofa.Count - 1].SetValue("Dagrofa", 4); // Leverandør
                                        listDagrofa[listDagrofa.Count - 1].SetValue(råvare[1], 5); // Råvare
                                        if (arrayDagrofa[row, 3].ToString() == "Ja") // konv/øko
                                        {
                                            listDagrofa[listDagrofa.Count - 1].SetValue("Øko", 6);
                                        }
                                        else
                                        {
                                            listDagrofa[listDagrofa.Count - 1].SetValue("Konv", 6);
                                        }

                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[row, 2].ToString(), 7); // Variant
                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[row, col + 3].ToString(), 8); // pris pr. enhed
                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[row, col + 2].ToString(), 9); // pris i alt
                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[row, col + 1].ToString(), 10); // Kg

                                        float totalPrice = float.Parse(listDagrofa[listDagrofa.Count - 1].GetValue(9).ToString());
                                        float totalWeight = float.Parse(listDagrofa[listDagrofa.Count - 1].GetValue(10).ToString());

                                        listDagrofa[listDagrofa.Count - 1].SetValue(totalPrice / totalWeight + "", 11); // kilopris
                                        listDagrofa[listDagrofa.Count - 1].SetValue(arrayDagrofa[row, 4].ToString(), 12); // oprindelse
                                    }
                                }
                                catch (NullReferenceException)
                                {

                                }
                            }
                        }
                        rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listDagrofa.Count));

                        object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                        // Sets the values in the Grønt Grossisten Object Array
                        for (int row = 0; row < listDagrofa.Count; row++)
                        {
                            for (int col = 0; col < 13; col++)
                            {
                                arrayMerged[row + 1, col + 1] = listDagrofa[row].GetValue(col);
                            }
                        }
                        rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                        rangeMerged = worksheetMerged.UsedRange;
                    }
                    catch
                    {

                    }
                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listDagrofa.Count)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listDagrofa.Count)).Font.Size = 11;

                    // Releasing the Excel interop objects
                    workbookDagrofa.Close(false);
                    MRCO(workbookDagrofa);
                    MRCO(worksheetDagrofa);
                    MRCO(rangeDagrofa);
                }

                // rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                // rangeMerged = worksheetMerged.UsedRange;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Code for Emmerys files
        private void ButtonEmmerysPath_Click(object sender, EventArgs e)
        {
            this.openEmmerysPathDialog.Title = "Select Emmerys File";
            if (openEmmerysPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathEmmerys = openEmmerysPathDialog.FileName;
                ButtonEmmerysPath.Text = openEmmerysPathDialog.FileName;
            }
        }
        private void EmmerysButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookEmmerys;
            Excel.Range rangeEmmerys;

            try
            {
                workbookEmmerys = excelProgram.Workbooks.Open(pathEmmerys);

                bool skipSheet = true;

                foreach (Excel.Worksheet worksheetEmmerys in workbookEmmerys.Sheets)
                {
                    if (!skipSheet)
                    {
                        rangeEmmerys = worksheetEmmerys.UsedRange;

                        int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                        int rowCountEmmerys = rangeEmmerys.Rows.Count;
                        int colCountEmmerys = rangeEmmerys.Columns.Count;

                        int headerRows = 10;

                        String infostringEmmerys = worksheetEmmerys.Cells[2, 2].Text;
                        String[] dateEmmerysString = infostringEmmerys.Split(new string[] { ".." }, StringSplitOptions.None);
                        DateTime dateEmmerys = DateTime.Parse(dateEmmerysString[1]);

                        // Imports the cell data from the Emmerys sheet as an array of Objects
                        Object[,] arrayEmmerys = rangeEmmerys.get_Value();

                        string currentHospital = worksheetEmmerys.Cells[1, 2].Text;

                        // Creates a List of String arrays for every rowOld in the BC worksheet.
                        // Amount of rows as a List to allow for deletion of irrelevant entries.
                        List<String[]> listEmmerys = new List<String[]>();

                        // For every row in the imported Hørkram Object array, copy its value to the corresponding String in the List of String arrays
                        int rowNew = 0;
                        for (int rowEmmerys = headerRows; rowEmmerys < rowCountEmmerys; rowEmmerys++)
                        {
                            if (!arrayEmmerys[rowEmmerys, 1].ToString().Contains("Produktnavn"))
                            {
                                listEmmerys.Add(new string[14]);
                                for (int col = 0; col < colCountEmmerys; col++)
                                {
                                    Debug.WriteLine(col);
                                    try // "Try" because the cell's value can be Null
                                    {
                                        listEmmerys[rowNew].SetValue(arrayEmmerys[rowEmmerys, col + 1].ToString(), col);
                                    }
                                    catch (NullReferenceException) // "Catch" in case the cell's value is Null
                                    {
                                        listEmmerys[rowNew].SetValue("", col);
                                    }
                                }
                                rowNew++;
                            }
                        }

                        rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listEmmerys.Count));

                        object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                        // Sets the values in the Emmerys Object Array
                        for (int row = 0; row < listEmmerys.Count; row++)
                        {
                            arrayMerged[row + 1, 1] = dateEmmerys.Year; // År
                            arrayMerged[row + 1, 2] = (dateEmmerys.Month) / 3; // Kvartal
                            arrayMerged[row + 1, 3] = currentHospital; // Hospital

                            string[] råvare = GetRåvare("Emmerys", listEmmerys[row].GetValue(0).ToString());

                            arrayMerged[row + 1, 4] = råvare[0]; // Råvarekategori
                            arrayMerged[row + 1, 5] = "Emmerys"; // Leverandør
                            arrayMerged[row + 1, 6] = råvare[1]; // Råvare
                            if (listEmmerys[row].GetValue(1) as String == "ØKO") // konv/øko
                            {
                                arrayMerged[row + 1, 7] = "Øko";
                            }
                            if (listEmmerys[row].GetValue(1) as String == "Konventionel")
                            {
                                arrayMerged[row + 1, 7] = "Konv";
                            }
                            arrayMerged[row + 1, 8] = listEmmerys[row].GetValue(0); // Varianter/opr
                            arrayMerged[row + 1, 9] = listEmmerys[row].GetValue(2); // Pris pr enhed
                            arrayMerged[row + 1, 10] = listEmmerys[row].GetValue(4); // Pris i alt
                            arrayMerged[row + 1, 11] = listEmmerys[row].GetValue(6); // Kg
                            arrayMerged[row + 1, 12] = float.Parse(listEmmerys[row].GetValue(4).ToString()) / float.Parse(listEmmerys[row].GetValue(6).ToString()); // Kilopris
                            arrayMerged[row + 1, 13] = "DAN"; // Oprindelse
                        }

                        rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                        rangeMerged = worksheetMerged.UsedRange;

                        //Format the cells.
                        worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listEmmerys.Count)).Font.Name = "Calibri";
                        worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listEmmerys.Count)).Font.Size = 11;

                        // Releasing the Excel interop objects for the worksheet
                        MRCO(rangeEmmerys);
                    }
                    skipSheet = false;
                    MRCO(worksheetEmmerys);
                }
                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects for workbook
                workbookEmmerys.Close(false);
                MRCO(workbookEmmerys);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Code for Frisksnit files
        private void ButtonFrisksnitPath_Click(object sender, EventArgs e)
        {
            this.openFrisksnitPathDialog.Multiselect = true;
            this.openFrisksnitPathDialog.Title = "Select Frisksnit files";

            if (openFrisksnitPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathFrisksnit = openFrisksnitPathDialog.FileNames.ToList();
                ButtonFrisksnitPath.Text = openFrisksnitPathDialog.FileName;
            }
        }
        private void FrisksnitButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookFrisksnit;
            Excel._Worksheet worksheetFrisksnit;
            Excel.Range rangeFrisksnit;

            try
            {
                foreach (String fileFrisksnit in pathFrisksnit)
                {
                    workbookFrisksnit = excelProgram.Workbooks.Open(fileFrisksnit);
                    worksheetFrisksnit = workbookFrisksnit.Sheets[1];
                    rangeFrisksnit = worksheetFrisksnit.UsedRange;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    int rowCountFrisksnit = rangeFrisksnit.Rows.Count;
                    int colCountFrisksnit = rangeFrisksnit.Columns.Count;

                    int headerRows = 7;
                    bool øko = false;

                    string currentHospital = worksheetFrisksnit.Cells[3, 1].Text;

                    String infostringFrisksnit = worksheetFrisksnit.Cells[4, 1].Text;
                    String[] dateFrisksnitString = infostringFrisksnit.Split(new string[] { " - " }, StringSplitOptions.None);
                    DateTime dateFrisksnit = DateTime.Parse(dateFrisksnitString[2]);

                    // Imports the cell data from the Frisksnit sheet as an array of Objects
                    Object[,] arrayFrisksnit = rangeFrisksnit.get_Value();

                    // Creates a List of String arrays for every row to be added to the merged worksheet.
                    // Amount of rows as a List to allow for deletion of irrelevant entries.
                    List<String[]> listFrisksnit = new List<String[]>();

                    // For every row in the imported Frisksnit Object array, copy its value to the corresponding String in the List of String arrays
                    int rowNew = 0;
                    for (int rowFrisksnit = headerRows; rowFrisksnit < rowCountFrisksnit - 1; rowFrisksnit++)
                    {
                        try
                        {
                            if (!(arrayFrisksnit[rowFrisksnit, 1].ToString().Contains("Total") | arrayFrisksnit[rowFrisksnit, 1].ToString().Contains("Gruppe")))
                            {
                                listFrisksnit.Add(new string[14]);
                                for (int col = 0; col < colCountFrisksnit; col++)
                                {
                                    Debug.WriteLine(col);
                                    try // "Try" because the cell's value can be Null
                                    {
                                        listFrisksnit[rowNew].SetValue(arrayFrisksnit[rowFrisksnit, col + 1].ToString(), col);
                                    }
                                    catch (NullReferenceException) // "Catch" in case the cell's value is Null
                                    {
                                        listFrisksnit[rowNew].SetValue("", col);
                                    }
                                }
                                rowNew++;
                            }
                        }
                        catch
                        {
                        }
                    }

                    rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listFrisksnit.Count));

                    object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    // Sets the values in the Frisksnit Object Array
                    for (int row = 0; row < listFrisksnit.Count; row++)
                    {
                        øko = false;
                        if (listFrisksnit[row].GetValue(0).ToString().Contains("Øko"))
                        {
                            øko = true;
                        }
                        arrayMerged[row + 1, 1] = dateFrisksnit.Year; // År
                        arrayMerged[row + 1, 2] = (dateFrisksnit.Month) / 3; // Kvartal
                        arrayMerged[row + 1, 3] = currentHospital; // Hospital
                        arrayMerged[row + 1, 4] = ""; // Råvarekategori
                        arrayMerged[row + 1, 5] = "Frisksnit"; // Leverandør
                        if (øko) // konv/øko
                        {
                            arrayMerged[row + 1, 6] = listFrisksnit[row].GetValue(0).ToString().Replace("Øko - ", ""); // Råvare
                        }
                        else
                        {
                            arrayMerged[row + 1, 6] = listFrisksnit[row].GetValue(0); // Råvare
                        }
                        if (øko) // konv/øko
                        {
                            arrayMerged[row + 1, 7] = "Øko";
                        }
                        else
                        {
                            arrayMerged[row + 1, 7] = "Konv";
                        }
                        arrayMerged[row + 1, 8] = listFrisksnit[row].GetValue(2); // Varianter/opr
                        arrayMerged[row + 1, 9] = float.Parse(listFrisksnit[row].GetValue(5).ToString()) / float.Parse(listFrisksnit[row].GetValue(3).ToString()); // Pris pr enhed
                        arrayMerged[row + 1, 10] = listFrisksnit[row].GetValue(5); // Pris i alt
                        arrayMerged[row + 1, 11] = listFrisksnit[row].GetValue(4); // Kg
                        arrayMerged[row + 1, 12] = float.Parse(listFrisksnit[row].GetValue(5).ToString()) / float.Parse(listFrisksnit[row].GetValue(4).ToString()); // Kilopris
                        arrayMerged[row + 1, 13] = "DAN"; // Oprindelse
                    }

                    rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                    rangeMerged = worksheetMerged.UsedRange;

                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listFrisksnit.Count)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listFrisksnit.Count)).Font.Size = 11;

                    //AutoFit columns A:V.
                    rangeMerged = worksheetMerged.get_Range("A1", "M1");
                    rangeMerged.EntireColumn.AutoFit();


                    //Make sure Excel is visible and give the user control
                    //of Microsoft Excel's lifetime.
                    excelProgram.Visible = true;
                    excelProgram.UserControl = true;

                    // Releasing the Excel interop objects
                    workbookFrisksnit.Close(false);
                    MRCO(workbookFrisksnit);
                    MRCO(worksheetFrisksnit);
                    MRCO(rangeFrisksnit);
                }
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Code for Grønt Grossisten files
        private void ButtonGrøntGrossistenPath_Click(object sender, EventArgs e)
        {
            this.openGrøntGrossistenPathDialog.Title = "Select Grønt Grossisten file";
            if (openGrøntGrossistenPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathGrøntGrossisten = openGrøntGrossistenPathDialog.FileName;
                ButtonGrøntGrossistenPath.Text = openGrøntGrossistenPathDialog.FileName;
            }
        }
        private void GrøntgrossistenButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookGrøntGrossisten;
            Excel._Worksheet worksheetGrøntGrossisten;
            Excel.Range rangeGrøntGrossisten;

            try
            {
                workbookGrøntGrossisten = excelProgram.Workbooks.Open(pathGrøntGrossisten);
                worksheetGrøntGrossisten = workbookGrøntGrossisten.Sheets[1];
                rangeGrøntGrossisten = worksheetGrøntGrossisten.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountGrøntGrossisten = rangeGrøntGrossisten.Rows.Count;
                int colCountGrøntGrossisten = rangeGrøntGrossisten.Columns.Count;

                // Imports the cell data from the Grønt Grossisten sheet as an array of Objects
                Object[,] arrayGrøntGrossisten = rangeGrøntGrossisten.get_Value();

                // Creates a List of String arrays for every rowOld in the Grønt Grossisten worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<String[]> listGrøntGrossisten = new List<String[]>();

                string currentHospital = "";

                int entries = 0;

                for (int row = 0; row < rowCountGrøntGrossisten; row++)
                {
                    try
                    {
                        if (IsNumeric(arrayGrøntGrossisten[row + 1, 1].ToString()))
                        {
                            listGrøntGrossisten.Add(new string[13]);
                            listGrøntGrossisten[entries].SetValue("Ikke oplyst", 0); // År
                            listGrøntGrossisten[entries].SetValue("Ikke oplyst", 1); // Kvartal
                            listGrøntGrossisten[entries].SetValue(currentHospital, 2); // Hospital
                            listGrøntGrossisten[entries].SetValue("Ikke oplyst", 3); // Råvarekategori
                            listGrøntGrossisten[entries].SetValue("Grønt Grossisten", 4); // Leverandør
                            listGrøntGrossisten[entries].SetValue("Ikke oplyst", 5); // Råvare
                            if (arrayGrøntGrossisten[row + 1, 9].ToString() == "1") // konv/øko
                            {
                                listGrøntGrossisten[entries].SetValue("Øko", 6);
                            }
                            else
                            {
                                listGrøntGrossisten[entries].SetValue("Konv", 6);
                            }
                            listGrøntGrossisten[entries].SetValue(arrayGrøntGrossisten[row + 1, 3].ToString(), 7); // Variant

                            string priceprunit = arrayGrøntGrossisten[row + 1, 12].ToString();
                            string pricetotalstring = arrayGrøntGrossisten[row + 1, 13].ToString();
                            string weighttotalstring = arrayGrøntGrossisten[row + 1, 8].ToString();

                            float pricetotalfloat = float.Parse(pricetotalstring);
                            float weighttotalfloat = float.Parse(weighttotalstring) / 1000;

                            listGrøntGrossisten[entries].SetValue(priceprunit, 8); // pris pr. enhed
                            listGrøntGrossisten[entries].SetValue(pricetotalstring, 9); // pris i alt
                            listGrøntGrossisten[entries].SetValue(weighttotalfloat + "", 10); // Kg
                            listGrøntGrossisten[entries].SetValue((pricetotalfloat / weighttotalfloat) + "", 11); // kilopris
                            listGrøntGrossisten[entries].SetValue(arrayGrøntGrossisten[row + 1, 10], 12); // oprindelse
                            entries++;
                        }
                        else
                        {
                            currentHospital = arrayGrøntGrossisten[row + 1, 1].ToString();
                            row++;
                        }
                    }
                    catch
                    {
                        row++;
                    }
                }

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listGrøntGrossisten.Count));

                object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                // Sets the values in the Grønt Grossisten Object Array
                for (int row = 0; row < listGrøntGrossisten.Count; row++)
                {
                    for (int col = 0; col < 13; col++)
                    {
                        arrayMerged[row + 1, col + 1] = listGrøntGrossisten[row].GetValue(col);
                    }
                }

                rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listGrøntGrossisten.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listGrøntGrossisten.Count)).Font.Size = 11;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects
                workbookGrøntGrossisten.Close(false);
                MRCO(workbookGrøntGrossisten);
                MRCO(worksheetGrøntGrossisten);
                MRCO(rangeGrøntGrossisten);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Code for Hørkram files
        private void ButtonHørkramPath_Click(object sender, EventArgs e)
        {
            this.openHørkramPathDialog.Title = "Select Hørkram file";
            if (openHørkramPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathHørkram = openHørkramPathDialog.FileName;
                ButtonHørkramPath.Text = openHørkramPathDialog.FileName;
            }
        }
        private void HørkramButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookHørkram;
            Excel._Worksheet worksheetHørkram;
            Excel._Worksheet infosheetHørkram;
            Excel.Range rangeHørkram;

            try
            {
                workbookHørkram = excelProgram.Workbooks.Open(pathHørkram);
                worksheetHørkram = workbookHørkram.Sheets[2];
                infosheetHørkram = workbookHørkram.Sheets[1];
                rangeHørkram = worksheetHørkram.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountHørkram = rangeHørkram.Rows.Count;
                int colCountHørkram = rangeHørkram.Columns.Count;

                // Imports the date from the Hørkram worksheet
                DateTime dateHørkram = DateTime.Parse(infosheetHørkram.Cells[5, 4].Text);

                // Imports the cell data from the Hørkram sheet as an array of Objects
                Object[,] arrayHørkram = rangeHørkram.get_Value();

                // Creates a List of String arrays for every rowOld in the Hørkram worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<String[]> listHørkram = new List<String[]>();

                // For every row in the imported Hørkram Object array, copy its value to the corresponding String in the List of String arrays
                for (int row = 0; row < rowCountHørkram; row++)
                {
                    listHørkram.Add(new string[14]);
                    for (int col = 0; col < colCountHørkram; col++)
                    {
                        try // "Try" because the cell's value can be Null
                        {
                            listHørkram[row].SetValue(arrayHørkram[row + 1, col + 1].ToString(), col);
                        }
                        catch (NullReferenceException) // "Catch" in case the cell's value is Null
                        {
                            listHørkram[row].SetValue("", col);
                        }
                    }
                }

                // Deletion of irrelevant entries from the List of String arrays
                listHørkram.RemoveRange(0, 2); // Header entries in rowOld 1 and 2
                listHørkram.RemoveAll(s => s[4].Contains("Non food") // Entries for non-food items
                || s[4].Contains("Hjælpevarenumre")
                || s[4].Contains("Engangsmateriale")
                || s[4].Contains("storkøkkentilbehør"));

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listHørkram.Count));

                object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                // Sets the values in the Hørkram Object Array
                for (int row = 0; row < listHørkram.Count; row++)
                {
                    arrayMerged[row + 1, 1] = dateHørkram.Year; // År
                    arrayMerged[row + 1, 2] = (dateHørkram.Month) / 3; // Kvartal
                    arrayMerged[row + 1, 3] = listHørkram[row].GetValue(1); // Hospital
                    arrayMerged[row + 1, 4] = listHørkram[row].GetValue(4); // Råvarekategori
                    arrayMerged[row + 1, 5] = "Hørkram"; // Leverandør
                    arrayMerged[row + 1, 6] = listHørkram[row].GetValue(5); // Råvare
                    if (listHørkram[row].GetValue(6) as String == "J") // konv/øko
                    {
                        arrayMerged[row + 1, 7] = "Øko";
                    }
                    if (listHørkram[row].GetValue(6) as String == "N")
                    {
                        arrayMerged[row + 1, 7] = "Konv";
                    }
                    arrayMerged[row + 1, 8] = listHørkram[row].GetValue(3); // Varianter/opr
                    arrayMerged[row + 1, 9] = float.Parse(listHørkram[row].GetValue(10) as String) / float.Parse(listHørkram[row].GetValue(9) as String); // Pris pr enhed
                    arrayMerged[row + 1, 10] = listHørkram[row].GetValue(10); // Pris i alt
                    arrayMerged[row + 1, 11] = listHørkram[row].GetValue(11); // Kg
                    arrayMerged[row + 1, 12] = listHørkram[row].GetValue(13); // Kilopris
                    arrayMerged[row + 1, 13] = listHørkram[row].GetValue(7); // Oprindelse
                }

                rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listHørkram.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listHørkram.Count)).Font.Size = 11;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects
                workbookHørkram.Close(false);
                MRCO(workbookHørkram);
                MRCO(worksheetHørkram);
                MRCO(infosheetHørkram);
                MRCO(rangeHørkram);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Path buttons code
        private void ButtonDeViKasPath_Click(object sender, EventArgs e)
        {
            this.openPathDialog.Multiselect = true;
            this.openPathDialog.Title = "Select DeViKas files";

            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathDeViKas = openPathDialog.FileNames.ToList();
                ButtonDeViKasPath.Text = openPathDialog.FileName;
            }
        }

        // Import buttons code
        private void ACButton_Click(object sender, System.EventArgs e)
        {
            _Workbook workbookSource;
            _Worksheet worksheetSource;
            Range rangeSource;

            try
            {
                foreach (string fileAC in pathAC)
                {
                    workbookSource = excelProgram.Workbooks.Open(fileAC);
                    worksheetSource = workbookSource.Sheets[1];
                    rangeSource = worksheetSource.UsedRange;

                    object[,] firstColumn = rangeSource.get_Value();

                    Debug.WriteLine(fileAC);

                    int rowCountSource = 2;
                    for (int count = 2; count < rangeSource.Rows.Count; count++)
                    {
                        try // 
                        {
                            if (float.Parse(firstColumn[count, 1].ToString()) > 0)
                            {
                                rowCountSource++;
                            }
                        }
                        catch (NullReferenceException) // "Catch" in case the cell's value is Null
                        {
                            break;
                        }
                    }
                    rangeSource = worksheetSource.get_Range("A1", "H" + rowCountSource);

                    int colCountAC = rangeSource.Columns.Count;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    // Imports the cell data from the AC sheet as an array of Objects
                    object[,] arrayImported = rangeSource.get_Value();

                    // Creates a List of String arrays for every rowOld in the AC worksheet.
                    // Amount of rows as a List to allow for deletion of irrelevant entries.
                    List<Row> listConverted = ConvertAC(arrayImported, rowCountSource);

                    rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                    object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                    arrayMerged = ConvertList(listConverted, arrayMerged);

                    rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                    rangeMerged = worksheetMerged.UsedRange;

                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count - 1)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count - 1)).Font.Size = 11;

                    //AutoFit columns A:V.
                    rangeMerged = worksheetMerged.get_Range("A1", "M1");
                    rangeMerged.EntireColumn.AutoFit();

                    //Make sure Excel is visible and give the user control
                    //of Microsoft Excel's lifetime.
                    excelProgram.Visible = true;
                    excelProgram.UserControl = true;

                    // Releasing the Excel interop objects
                    workbookSource.Close(false);
                    MRCO(workbookSource);
                    MRCO(worksheetSource);
                    MRCO(rangeSource);
                }
            }
            catch (Exception theException)
            {
                string errorMessage;
                errorMessage = "Error: ";
                errorMessage = string.Concat(errorMessage, theException.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
        private void BCButton_Click(object sender, System.EventArgs e)
        {
            _Workbook workbookSource;
            _Worksheet worksheetSource;
            _Worksheet infosheetSource;
            Range rangeSource;
            Range rangeInfo;

            try
            {
                workbookSource = excelProgram.Workbooks.Open(pathBC);
                worksheetSource = workbookSource.Sheets[1];
                infosheetSource = workbookSource.Sheets[2];
                rangeSource = worksheetSource.UsedRange;
                rangeInfo = infosheetSource.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountSource = rangeSource.Rows.Count;
                int colCountSource = rangeSource.Columns.Count;

                // Imports the cell data from the Hørkram sheet as an array of Objects
                object[,] arraySource = rangeSource.get_Value();
                object[,] arrayInfo = rangeInfo.get_Value();

                // Creates a List of String arrays for every rowOld in the BC worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<Row> listConverted = ConvertBC(arraySource, rowCountSource, arrayInfo);

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (listConverted.Count + usedRowsMerged));
                object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                arrayMerged = ConvertList(listConverted, arrayMerged);

                rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (listConverted.Count + usedRowsMerged)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (listConverted.Count + usedRowsMerged)).Font.Size = 11;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();


                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects
                workbookSource.Close(false);
                MRCO(workbookSource);
                MRCO(worksheetSource);
                MRCO(infosheetSource);
                MRCO(rangeSource);

            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = string.Concat(errorMessage, theException.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
        private void DeViKasButton_Click(object sender, EventArgs e)
        {
            _Workbook workbookSource;
            Range rangeSource;
            try
            {
                foreach (string fileDeViKas in pathDeViKas)
                {
                    workbookSource = excelProgram.Workbooks.Open(fileDeViKas);
                    foreach (Worksheet worksheetDeViKas in workbookSource.Sheets)
                    {
                        rangeSource = worksheetDeViKas.UsedRange;

                        int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                        int rowCountSource = rangeSource.Rows.Count;
                        int colCountSource = rangeSource.Columns.Count;

                        object[,] arrayImported = rangeSource.get_Value();

                        bool splitWeight = false;
                        if (arrayImported[12, 7].ToString().Contains("Øko"))
                        {
                            splitWeight = true;
                        }

                        List<Row> listConverted = ConvertDeViKas(arrayImported, splitWeight, rowCountSource);

                        if(listConverted.Count > 0)
                        {
                            rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                            object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                            arrayMerged = ConvertList(listConverted, arrayMerged);

                            rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                            rangeMerged = worksheetMerged.UsedRange;

                            //Format the cells.
                            worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                            worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;
                        }
                        // Releasing the Excel interop objects for the worksheet
                        MRCO(rangeSource);
                        MRCO(worksheetDeViKas);
                    }

                    // Releasing the Excel interop objects for workbook
                    workbookSource.Close(false);
                    MRCO(workbookSource);
                }

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;
            }
            catch (Exception theException)
            {
                string errorMessage;
                errorMessage = "Error: ";
                errorMessage = string.Concat(errorMessage, theException.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        // Support code
        public void MRCO(object comObject) // Based on code from breezetree.com/blog/
        {
            if (comObject != null)
            {
                Marshal.ReleaseComObject(comObject);
                comObject = null;
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            MRCO(excelProgram);
            MRCO(workbookMerged);
            MRCO(worksheetMerged);
            MRCO(rangeMerged);
        }

        private string[] GetRåvare(string company, string variant)
        {
            List<String[]> listCompany = listLibrary.Where(x => x[1] == company).ToList();

            List<String[]> listVariant = listCompany.Where(x => x[4] == variant).ToList();

            string[] categories = new string[2];

            if (listVariant.Count > 0)
            {
                categories[0] = listVariant[0].GetValue(0).ToString(); // Råvarekategori
                categories[1] = listVariant[0].GetValue(2).ToString(); // Råvare
            }
            else
            {
                categories[0] = "";
                categories[1] = "";
            }

            return categories;
        }
        private string GetRåvare(string company, string variant, bool getcategory)
        {
            List<String[]> listCompany = listLibrary.Where(x => x[1] == company).ToList();

            List<String[]> listVariant = listCompany.Where(x => x[4] == variant).ToList();

            string[] categories = new string[2];

            if (listVariant.Count > 0)
            {
                categories[0] = listVariant[0].GetValue(0).ToString(); // Råvarekategori
                categories[1] = listVariant[0].GetValue(2).ToString(); // Råvare
            }
            else
            {
                categories[0] = "";
                categories[1] = "";
            }
            if (getcategory)
            {
                return categories[0];
            }
            else
            {
                return categories[1];
            }
        }

        private bool IsNumeric(string input)
        {
            return float.TryParse(input, out _);
        }

        private string GetQuarter(string input)
        {
            return (int.Parse(input) + 2) / 3 + "";
        }

        // Transformations from input Array to standardised List of Rows
        internal List<Row> ConvertDeViKas(Object[,] inputMatrix, bool splitWeight, int rowCount)
        {
            string year;
            string month;
            string quarter;

            try
            {
                year = DateTime.Parse(inputMatrix[3, 2].ToString()).Year + "";
                month = DateTime.Parse(inputMatrix[3, 2].ToString()).Month + "";
                quarter = "K" + GetQuarter(month);
            }
            catch
            {
                year = inputMatrix[1, 1].ToString().Split(' ').Last();
                try
                {
                    month = inputMatrix[4, 4].ToString().Split('/','-','.').Last();
                    quarter = "K" + GetQuarter(month);
                }
                catch
                {
                    quarter = "Fejl i data";
                }
            }

            int headerRows = 13;

            List<Row> output = new List<Row>();

            if (splitWeight)
            {
                int rowOutput = 0;
                for (int rowInput = headerRows; rowInput < rowCount; rowInput++)
                {
                    try
                    {
                        int amount = int.Parse(inputMatrix[rowInput, 4].ToString());

                        if(amount > 0)
                        {
                            string ecology = "Konv";
                            string weight = "";

                            try
                            {
                                weight = float.Parse(inputMatrix[rowInput, 6].ToString()) / 1000 + "";
                            }
                            catch
                            {
                                weight = float.Parse(inputMatrix[rowInput, 7].ToString()) / 1000 + "";
                                ecology = "Øko";
                            }

                            Row newEntry = new Row(
                                år: year,
                                kvartal: quarter,
                                hospital: inputMatrix[8, 2].ToString(),
                                råvarekategori: "Bagværk / søde sager",
                                leverandør: "DeViKas Bageri",
                                råvare: "Bagværk",
                                øko: ecology,
                                variant: inputMatrix[rowInput, 1].ToString(),
                                prisEnhed: inputMatrix[rowInput, 3].ToString(),
                                prisTotal: inputMatrix[rowInput, 5].ToString(),
                                kg: weight,
                                oprindelse: "DAN"
                                );
                            output.Add(newEntry);
                            rowOutput++;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            else
            {
                int rowOutput = 0;
                for (int rowInput = headerRows; rowInput < rowCount; rowInput++)
                {
                    try
                    {
                        int amount = int.Parse(inputMatrix[rowInput, 5].ToString());

                        if (amount > 0)
                        {
                            string ecology = "Konv";

                            if (inputMatrix[rowInput, 1].ToString().Contains("økologisk"))
                            {
                                ecology = "Øko";
                            }

                            Row newEntry = new Row(
                                år: year,
                                kvartal: quarter,
                                hospital: inputMatrix[8, 2].ToString(),
                                råvarekategori: "Bagværk / søde sager",
                                leverandør: "DeViKas Bageri",
                                råvare: "Bagværk",
                                øko: ecology,
                                variant: inputMatrix[rowInput, 1].ToString(),
                                prisEnhed: inputMatrix[rowInput, 4].ToString(),
                                prisTotal: inputMatrix[rowInput, 6].ToString(),
                                kg: float.Parse(inputMatrix[rowInput, 7].ToString()) / 1000 + "",
                                oprindelse: "DAN"
                                );
                            output.Add(newEntry);
                            rowOutput++;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            return output;
        }
        internal List<Row> ConvertAC(object[,] inputMatrix, int rowCount)
        {
            string year = "2021";
            string month;
            string quarter = "K2";
            string ecology = "Konv";

            int headerrows = 2;

            List<Row> output = new List<Row>();

            for (int rowInput = headerrows; rowInput < rowCount; rowInput++)
            {
                if (inputMatrix[rowInput, 4].ToString().Contains("ØKO"))
                {
                    ecology = "Øko";
                }

                string oprindelse = inputMatrix[rowInput, 4].ToString().Replace(" ", "").Split('(').Last();

                Row newEntry = new Row(
                    år: year,
                    kvartal: quarter,
                    hospital: inputMatrix[rowInput, 2].ToString(),
                    råvarekategori: GetRåvare("AC", inputMatrix[rowInput, 4].ToString(), true),
                    leverandør: "AC",
                    råvare: GetRåvare("AC", inputMatrix[rowInput, 4].ToString(), false),
                    øko: ecology,
                    variant: inputMatrix[rowInput, 4].ToString(),
                    prisEnhed: inputMatrix[rowInput, 8].ToString(),
                    prisTotal: inputMatrix[rowInput, 7].ToString(),
                    kg: inputMatrix[rowInput, 6].ToString(),
                    oprindelse: oprindelse
                    );
                output.Add(newEntry);

            }
            return output;
        }
        internal List<Row> ConvertBC(object[,] inputMatrix, int rowCount, object[,] hospitalMatrix)
        {
            DateTime dateTime = DateTime.Parse(inputMatrix[2,2].ToString().Split(new string[] { ".." }, StringSplitOptions.None).Last());

            string year = dateTime.Year + "";
            string month = dateTime.Month + "";
            string quarter = "K" + GetQuarter(month);
            string ecology = "Konv";
            string weight;
            string currentHospital = "";

            bool categoryEnd = false;

            int headerrows = 6;
            int rowOutput = 0;

            List<Row> output = new List<Row>();
            List<string> hospitalList = new List<string>();

            foreach(object hospital in hospitalMatrix.GetColumn(0))
            {
                hospitalList.Add(hospital.ToString());
            }

            for (int rowInput = headerrows; rowInput < rowCount; rowInput++)
            {
                Debug.WriteLine(rowInput);
                categoryEnd = false;
                try
                {
                    if (inputMatrix[rowInput, 2].ToString().Length > 0)
                    {
                        if (hospitalList.Contains(inputMatrix[rowInput, 2].ToString()))
                        {
                            int hospitalIndex = hospitalList.FindIndex(a => a == inputMatrix[rowInput, 2].ToString());
                            currentHospital = hospitalMatrix[hospitalIndex+1, 2].ToString();
                        }
                    }
                }
                catch
                {
                }
                while (!categoryEnd)
                {
                    try
                    {
                        try
                        {
                            weight = inputMatrix[rowInput, 9].ToString();
                            ecology = "Øko";
                        }
                        catch
                        {
                            try
                            {
                                weight = inputMatrix[rowInput, 10].ToString();
                            }
                            catch
                            {
                                weight = inputMatrix[rowInput, 11].ToString();
                            }
                        }
                        Row newEntry = new Row(
                            år: year,
                            kvartal: quarter,
                            hospital: currentHospital,
                            råvarekategori: inputMatrix[rowInput, 19].ToString(),
                            leverandør: "BC",
                            råvare: inputMatrix[rowInput, 20].ToString(),
                            øko: ecology,
                            variant: inputMatrix[rowInput, 2].ToString(),
                            prisEnhed: inputMatrix[rowInput, 8].ToString(),
                            prisTotal: inputMatrix[rowInput, 7].ToString(),
                            kg: weight,
                            oprindelse: inputMatrix[rowInput, 17].ToString()
                            );
                        output.Add(newEntry);
                        rowOutput++;
                        Debug.WriteLine(rowInput);
                        rowInput++;
                    }
                    catch
                    {
                        categoryEnd = true;
                    }
                }
            }

            return output;
        }

        // Model code
        internal object[,] ConvertList(List<Row> inputList, object[,] outputMatrix)
        {
            Debug.WriteLine("ConvertList");
            Debug.WriteLine(outputMatrix.Length);
            // Sets the values in the Emmerys Object Array
            for (int row = 0; row < inputList.Count; row++)
            {
                Debug.WriteLine(row);
                outputMatrix[row + 1, 1] = inputList[row].år;
                outputMatrix[row + 1, 2] = inputList[row].kvartal;
                outputMatrix[row + 1, 3] = inputList[row].hospital; // Hospital
                outputMatrix[row + 1, 4] = inputList[row].råvarekategori; // Råvarekategori
                outputMatrix[row + 1, 5] = inputList[row].leverandør; // Leverandør
                outputMatrix[row + 1, 6] = inputList[row].råvare; // Råvare
                outputMatrix[row + 1, 7] = inputList[row].øko; // Konv/øko
                outputMatrix[row + 1, 8] = inputList[row].variant; // Varianter/opr
                outputMatrix[row + 1, 9] = inputList[row].prisEnhed; // Pris pr enhed
                outputMatrix[row + 1, 10] = inputList[row].prisTotal; // Pris i alt
                outputMatrix[row + 1, 11] = inputList[row].kg; // Kg
                outputMatrix[row + 1, 12] = inputList[row].kilopris; // Kilopris
                outputMatrix[row + 1, 13] = inputList[row].oprindelse; // Oprindelse
            }
            return outputMatrix;
        }
    }
}