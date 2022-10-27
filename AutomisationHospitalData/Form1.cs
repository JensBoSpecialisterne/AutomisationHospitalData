using CommunityToolkit.HighPerformance;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        List<string[]> listLibrary = new List<string[]>();

        private readonly OpenFileDialog openPathDialog = new OpenFileDialog();

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
        private void Form1_Closed(object sender, FormClosedEventArgs e)
        {
            MRCO(excelProgram);
            MRCO(workbookMerged);
            MRCO(worksheetMerged);
            MRCO(rangeMerged);
        }
        private void Form1_Closing(object sender, FormClosingEventArgs e)
        {
            MRCO(excelProgram);
            MRCO(workbookMerged);
            MRCO(worksheetMerged);
            MRCO(rangeMerged);
        }
        private void ButtonBibliotekPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select Bibliotek File";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathBibliotek = openPathDialog.FileName;
                ButtonBibliotekPath.Text = openPathDialog.FileName;
            }
        }
        private void CreateBibliotekButton_Click(object sender, EventArgs e)
        {
            _Workbook workbookLibrary;
            _Worksheet worksheetLibrary;
            Range rangeLibrary;

            workbookLibrary = excelProgram.Workbooks.Open(pathBibliotek);
            worksheetLibrary = workbookLibrary.Sheets[1];
            rangeLibrary = worksheetLibrary.UsedRange;

            int rowCountLibrary = rangeLibrary.Rows.Count;
            int colCountLibrary = rangeLibrary.Columns.Count;

            object[,] arrayLibrary = rangeLibrary.get_Value();

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

        // Code for Grønt Grossisten files

        // Code for Hørkram files
        private void HørkramButton_Click(object sender, EventArgs e)
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
        private void ButtonACPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Multiselect = true;
            openPathDialog.Title = "Select AC files";

            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathAC = openPathDialog.FileNames.ToList();
                ButtonACPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonBCPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select BC File";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathBC = openPathDialog.FileName;
                ButtonBCPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonCBPBageriPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select CBP Bageri File";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathCBPBageri = openPathDialog.FileName;
                ButtonCBPBageriPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonDagrofaPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Multiselect = true;
            openPathDialog.Title = "Select Dagrofa files";

            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathDagrofa = openPathDialog.FileNames.ToList();
                ButtonDagrofaPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonDeViKasPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Multiselect = true;
            openPathDialog.Title = "Select DeViKas files";

            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathDeViKas = openPathDialog.FileNames.ToList();
                ButtonDeViKasPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonEmmerysPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select Emmerys File";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathEmmerys = openPathDialog.FileName;
                ButtonEmmerysPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonFrisksnitPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Multiselect = true;
            openPathDialog.Title = "Select Frisksnit files";

            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathFrisksnit = openPathDialog.FileNames.ToList();
                ButtonFrisksnitPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonGrøntGrossistenPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select Grønt Grossisten file";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathGrøntGrossisten = openPathDialog.FileName;
                ButtonGrøntGrossistenPath.Text = openPathDialog.FileName;
            }
        }
        private void ButtonHørkramPath_Click(object sender, EventArgs e)
        {
            openPathDialog.Title = "Select Hørkram file";
            if (openPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathHørkram = openPathDialog.FileName;
                ButtonHørkramPath.Text = openPathDialog.FileName;
            }
        }

        // Import buttons code
        private void ACButton_Click(object sender, EventArgs e)
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
        private void BCButton_Click(object sender, EventArgs e)
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
        private void CBPBageriButton_Click(object sender, EventArgs e)
        {

            _Workbook workbookSource;
            _Worksheet worksheetSource;
            _Worksheet infosheetSource;
            Range rangeSource;
            Range rangeInfo;

            try
            {
                workbookSource = excelProgram.Workbooks.Open(pathCBPBageri);
                worksheetSource = workbookSource.Sheets[2];
                infosheetSource = workbookSource.Sheets[1];
                rangeSource = worksheetSource.UsedRange;
                rangeInfo = infosheetSource.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountSource = rangeSource.Rows.Count;
                int colCountSource = rangeSource.Columns.Count;

                object[,] arraySource = rangeSource.get_Value();
                object[,] arrayInfo = rangeInfo.get_Value();

                List<Row> listConverted = ConvertCBPBageri(arraySource, rowCountSource, arrayInfo);

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (listConverted.Count + usedRowsMerged));
                object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                arrayMerged = ConvertList(listConverted, arrayMerged);

                rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;

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
        private void DagrofaButton_Click(object sender, EventArgs e)
        {

            _Workbook workbookSource;
            _Worksheet worksheetSource;
            Range rangeSource;

            try
            {
                foreach (string fileDagrofa in pathDagrofa)
                {
                    workbookSource = excelProgram.Workbooks.Open(fileDagrofa);
                    worksheetSource = workbookSource.Sheets[1];
                    rangeSource = worksheetSource.UsedRange;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    int rowCountSource = rangeSource.Rows.Count;
                    int colCountSource = rangeSource.Columns.Count;

                    // Imports the cell data from the AC sheet as an array of Objects
                    object[,] arrayImported = rangeSource.get_Value();

                    // Creates a List of String arrays for every rowOld in the AC worksheet.
                    // Amount of rows as a List to allow for deletion of irrelevant entries.
                    List<Row> listConverted = ConvertDagrofa(arrayImported, rowCountSource, colCountSource);

                    rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                    object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                    arrayMerged = ConvertList(listConverted, arrayMerged);

                    rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                    rangeMerged = worksheetMerged.UsedRange;

                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;

                    // Releasing the Excel interop objects
                    workbookSource.Close(false);
                    MRCO(workbookSource);
                    MRCO(worksheetSource);
                    MRCO(rangeSource);
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
                string errorMessage;
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
        private void EmmerysButton_Click(object sender, EventArgs e)
        {

            _Workbook workbookSource;
            Range rangeSource;

            try
            {
                workbookSource = excelProgram.Workbooks.Open(pathEmmerys);

                bool skipSheet = true;

                foreach (Worksheet worksheetSource in workbookSource.Sheets)
                {
                    if (!skipSheet)
                    {
                        rangeSource = worksheetSource.UsedRange;

                        int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                        int rowCountSource = rangeSource.Rows.Count;
                        int colCountSource = rangeSource.Columns.Count;

                        // Imports the cell data from the Emmerys sheet as an array of Objects
                        object[,] arraySource = rangeSource.get_Value();

                        // Creates a List of String arrays for every rowOld in the BC worksheet.
                        // Amount of rows as a List to allow for deletion of irrelevant entries.
                        List<Row> listConverted = ConvertEmmerys(arraySource, rowCountSource);

                        rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                        object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                        arrayMerged = ConvertList(listConverted, arrayMerged);

                        rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                        rangeMerged = worksheetMerged.UsedRange;

                        //Format the cells.
                        worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                        worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;

                        // Releasing the Excel interop objects for the worksheet
                        MRCO(rangeSource);
                    }
                    skipSheet = false;
                    MRCO(worksheetSource);
                }
                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects for workbook
                workbookSource.Close(false);
                MRCO(workbookSource);
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
        private void FrisksnitButton_Click(object sender, EventArgs e)
        {

            _Workbook workbookSource;
            _Worksheet worksheetSource;
            Range rangeSource;

            try
            {
                foreach (string fileFrisksnit in pathFrisksnit)
                {
                    workbookSource = excelProgram.Workbooks.Open(fileFrisksnit);
                    worksheetSource = workbookSource.Sheets[1];
                    rangeSource = worksheetSource.UsedRange;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    int rowCountSource = rangeSource.Rows.Count;
                    int colCountSource = rangeSource.Columns.Count;

                    // Imports the cell data from the Frisksnit sheet as an array of Objects
                    object[,] arraySource = rangeSource.get_Value();

                    // Creates a List of String arrays for every row to be added to the merged worksheet.
                    // Amount of rows as a List to allow for deletion of irrelevant entries.
                    List<Row> listConverted = ConvertFrisksnit(arraySource, rowCountSource);

                    rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                    object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                    arrayMerged = ConvertList(listConverted, arrayMerged);

                    rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                    rangeMerged = worksheetMerged.UsedRange;

                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;

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
        private void GrøntgrossistenButton_Click(object sender, EventArgs e)
        {

            _Workbook workbookSource;
            _Worksheet worksheetSource;
            Range rangeSource;

            try
            {
                workbookSource = excelProgram.Workbooks.Open(pathGrøntGrossisten);
                worksheetSource = workbookSource.Sheets[1];
                rangeSource = worksheetSource.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountSource = rangeSource.Rows.Count;
                int colCountSource = rangeSource.Columns.Count;

                // Imports the cell data from the Grønt Grossisten sheet as an array of Objects
                object[,] arraySource = rangeSource.get_Value();

                // Creates a List of String arrays for every rowOld in the Grønt Grossisten worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<Row> listConverted = ConvertGrøntGrossisten(arraySource, rowCountSource);

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listConverted.Count));

                Debug.WriteLine("Next");
                object[,] arrayMerged = rangeMerged.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                arrayMerged = ConvertList(listConverted, arrayMerged);

                rangeMerged.set_Value(XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listConverted.Count)).Font.Size = 11;

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
        private string GetRåvare(string company, string variant, bool getcategory)
        {
            List<string[]> listCompany = listLibrary.Where(x => x[1] == company).ToList();

            List<string[]> listVariant = listCompany.Where(x => x[4] == variant).ToList();

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
        internal List<Row> ConvertAC(object[,] inputMatrix, int rowCount)
        {
            string year = "2021";
            string quarter = "K2";
            string ecology;

            int headerrows = 2;

            List<Row> output = new List<Row>();

            for (int rowInput = headerrows; rowInput <= rowCount; rowInput++)
            {
                Debug.WriteLine(rowInput);
                try
                {
                    ecology = "Konv";
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
                catch { }

            }
            return output;
        }
        internal List<Row> ConvertBC(object[,] inputMatrix, int rowCount, object[,] hospitalMatrix)
        {
            DateTime dateTime = DateTime.Parse(inputMatrix[2,2].ToString().Split(new string[] { ".." }, StringSplitOptions.None).Last());

            string year = dateTime.Year + "";
            string month = dateTime.Month + "";
            string quarter = "K" + GetQuarter(month);
            string ecology;
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

            for (int rowInput = headerrows; rowInput <= rowCount; rowInput++)
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
                catch { }
                while (!categoryEnd)
                {
                    try
                    {
                        ecology = "Konv";
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
        internal List<Row> ConvertCBPBageri(object[,] inputMatrix, int rowCount, object[,] infoMatrix)
        {
            string quarter = "K" + infoMatrix[1,2].ToString().Substring(0, 1);
            string ecology;

            int headerrows = 2;

            List<Row> output = new List<Row>();

            for(int rowInput = headerrows; rowInput <= rowCount; rowInput++)
            {
                ecology = "Øko";
                try
                {
                    if (inputMatrix[rowInput, 10].ToString().Contains("Ej"))
                    {
                        ecology = "Konv";
                    }

                    string variant = inputMatrix[rowInput, 2].ToString().Split(new string[] { " ~ " }, StringSplitOptions.None).Last();

                    if (float.Parse(inputMatrix[rowInput, 3].ToString()) > 0 &
                        !(
                        variant.Contains("gangshue") | 
                        variant.Contains("bæger") | 
                        variant.Contains("pose") | 
                        variant.Contains("låg") | 
                        variant.Contains("Palle") | 
                        variant.Contains("papir") | 
                        variant.Contains("Credi")
                        ))
                    {
                        Row newEntry = new Row(
                            år: inputMatrix[rowInput, 10].ToString(),
                            kvartal: quarter,
                            hospital: inputMatrix[rowInput, 1].ToString().Split(new string[] { " ~ " }, StringSplitOptions.None).Last(),
                            råvarekategori: GetRåvare("CBP", variant, true),
                            leverandør: "CBP Bageri",
                            råvare: GetRåvare("CBP", variant, false),
                            øko: ecology,
                            variant: variant,
                            prisEnhed: "",
                            prisTotal: inputMatrix[rowInput, 4].ToString(),
                            kg: inputMatrix[rowInput, 3].ToString(),
                            oprindelse: "DAN"
                            );
                        output.Add(newEntry);
                    }
                }
                catch { }
            }
            return output;
        }
        internal List<Row> ConvertDagrofa(object[,] inputMatrix, int rowCount, int colCount)
        {
            string[] headerInfo = inputMatrix[1, 1].ToString().Split(' ');

            string year = headerInfo.Last();
            string quarter = headerInfo[headerInfo.Length - 2].Replace('Q','K');
            string ecology;
            string variant;
            string origin;

            int headerRows = 9;
            int headerCols = 5;

            List<Row> output = new List<Row>();

            for (int rowInput = headerRows; rowInput < rowCount-1; rowInput++)
            {
                Debug.WriteLine("Row" + rowInput);
                ecology = "Konv";
                if (inputMatrix[rowInput, 3].ToString().Contains("Ja"))
                {
                    ecology = "Øko";
                }
                variant = inputMatrix[rowInput, 2].ToString();
                origin = inputMatrix[rowInput, 4].ToString();
                if (!(variant.Contains("Pant") |
                    variant.Contains("glas") |
                    variant.Contains("Låg") |
                    variant.Contains("låg") |
                    variant.Contains("bøtte") |
                    variant.Contains("Levering") |
                    variant.Contains("levering") |
                    variant.Contains("Gebyr") |
                    variant.Contains("gebyr") |
                    variant.Contains("smuld")
                    ))
                {
                    for (int colInput = headerCols; colInput < colCount; colInput += 4)
                    {
                        Debug.WriteLine("Row" + rowInput + " Col" + colInput);

                        string hospital = "";
                        string[] hospitalArray = inputMatrix[7, colInput].ToString().Split(' ');
                        for (int i = 1; i < hospitalArray.Length; i++)
                        {
                            hospital = hospital + hospitalArray[i] + " ";
                        }


                        try
                        {
                            if (float.Parse(inputMatrix[rowInput, colInput].ToString()) > 0)
                            {
                                Row newEntry = new Row(
                                    år: year,
                                    kvartal: quarter,
                                    hospital: hospital,
                                    råvarekategori: GetRåvare("Dagrofa", variant, true),
                                    leverandør: "Dagrofa",
                                    råvare: GetRåvare("Dagrofa", variant, false),
                                    øko: ecology,
                                    variant: variant,
                                    prisEnhed: inputMatrix[rowInput, colInput + 3].ToString(),
                                    prisTotal: inputMatrix[rowInput, colInput + 2].ToString(),
                                    kg: inputMatrix[rowInput, colInput + 1].ToString(),
                                    oprindelse: origin
                                    );
                                output.Add(newEntry);
                            }
                        }
                        catch { }
                    }
                }
            }
            return output;
        }
        internal List<Row> ConvertDeViKas(object[,] inputMatrix, bool splitWeight, int rowCount)
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
                    month = inputMatrix[4, 4].ToString().Split('/', '-', '.').Last();
                    quarter = "K" + GetQuarter(month);
                }
                catch
                {
                    quarter = "";
                }
            }

            int headerRows = 13;

            List<Row> output = new List<Row>();

            if (splitWeight)
            {
                int rowOutput = 0;
                for (int rowInput = headerRows; rowInput <= rowCount; rowInput++)
                {
                    try
                    {
                        int amount = int.Parse(inputMatrix[rowInput, 4].ToString());

                        if (amount > 0)
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
                    catch { }
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
                    catch { }
                }
            }
            return output;
        }
        internal List<Row> ConvertEmmerys(object[,] inputMatrix, int rowCount)
        {
            DateTime dateTime = DateTime.Parse(inputMatrix[2,2].ToString().Split(new string[] { ".." }, StringSplitOptions.None).Last());
            string year = dateTime.Year + "";
            string month = dateTime.Month + "";
            string quarter = "K" + GetQuarter(month);
            string hospital = inputMatrix[1, 2].ToString();
            string ecology;

            int headerrows;

            for(headerrows = 1; headerrows < rowCount; headerrows++)
            {
                try
                {
                    if(inputMatrix[headerrows, 1].ToString().Contains("Produktnavn"))
                    {
                        headerrows++;
                        break;
                    }
                }
                catch{}
            }

            List<Row> output = new List<Row>();

            for (int rowInput = headerrows; rowInput <= rowCount; rowInput++)
            {
                try
                {
                    ecology = "Konv";
                    if (inputMatrix[rowInput, 2].ToString().Contains("ØKO"))
                    {
                        ecology = "Øko";
                    }

                    string variant = inputMatrix[rowInput, 1].ToString();

                    Row newEntry = new Row(
                        år: year,
                        kvartal: quarter,
                        hospital: hospital,
                        råvarekategori: GetRåvare("Emmerys", variant, true),
                        leverandør: "Emmerys",
                        råvare: GetRåvare("Emmerys", variant, false),
                        øko: ecology,
                        variant: variant,
                        prisEnhed: inputMatrix[rowInput, 3].ToString(),
                        prisTotal: inputMatrix[rowInput, 5].ToString(),
                        kg: inputMatrix[rowInput, 7].ToString(),
                        oprindelse: "DAN"
                        );
                    output.Add(newEntry);
                }
                catch { }
            }
            return output;
        }
        internal List<Row> ConvertFrisksnit(object[,] inputMatrix, int rowCount)
        {
            string[] dateArray = inputMatrix[3, 1].ToString().Split(new string[] { " - " }, StringSplitOptions.None);
            DateTime dateTime = DateTime.Parse(dateArray[dateArray.Length-2]);

            string year = dateTime.Year + "";
            string month = dateTime.Month + "";
            string quarter = "K" + GetQuarter(month);
            string ecology;
            string hospital = inputMatrix[2,1].ToString();

            int headerrows = 5;

            List<Row> output = new List<Row>();

            for (int rowInput = headerrows; rowInput <= rowCount; rowInput++)
            {
                try
                {
                    string variant = inputMatrix[rowInput, 3].ToString();
                    float priceTotal = float.Parse(inputMatrix[rowInput, 6].ToString());
                    float weight = float.Parse(inputMatrix[rowInput, 5].ToString());
                    float amount = float.Parse(inputMatrix[rowInput, 4].ToString());

                    ecology = "Konv";
                    if (inputMatrix[rowInput, 3].ToString().Contains("Økologisk"))
                    {
                        ecology = "Øko";
                    }

                    Row newEntry = new Row(
                        år: year,
                        kvartal: quarter,
                        hospital: hospital,
                        råvarekategori: GetRåvare("Frisksnit", variant, true),
                        leverandør: "Frisksnit",
                        råvare: GetRåvare("Frisksnit", variant, false),
                        øko: ecology,
                        variant: variant,
                        prisEnhed: priceTotal / amount + "",
                        prisTotal: priceTotal + "",
                        kg: weight + "",
                        oprindelse: ""
                        );
                    output.Add(newEntry);
                }
                catch { }
            }
            return output;
        }
        internal List<Row> ConvertGrøntGrossisten(object[,] inputMatrix, int rowCount)
        {

            string year = "2021";
            string quarter = "K2";
            string ecology;
            string hospital = "";
            bool skip;

            List<Row> output = new List<Row>();

            for (int rowInput = 0; rowInput <= rowCount; rowInput++)
            {
                try
                {
                    skip = false;
                    if (inputMatrix[rowInput,1].ToString().Contains("Source No_"))
                    {
                        hospital = inputMatrix[rowInput-1,1].ToString();
                        skip = true;
                    }
                    if (!skip)
                    {
                        string variant = inputMatrix[rowInput, 3].ToString();
                        float priceTotal = float.Parse(inputMatrix[rowInput, 13].ToString());
                        float weight = float.Parse(inputMatrix[rowInput, 8].ToString())/1000;
                        float amount = float.Parse(inputMatrix[rowInput, 6].ToString());

                        ecology = "Konv";
                        if (inputMatrix[rowInput, 3].ToString().Contains("øko"))
                        {
                            ecology = "Øko";
                        }

                        Row newEntry = new Row(
                            år: year,
                            kvartal: quarter,
                            hospital: hospital,
                            råvarekategori: GetRåvare("Grønt Grossisten", variant, true),
                            leverandør: "Grønt Grossisten",
                            råvare: GetRåvare("Grønt Grossisten", variant, false),
                            øko: ecology,
                            variant: variant,
                            prisEnhed: inputMatrix[rowInput, 12].ToString(),
                            prisTotal: priceTotal + "",
                            kg: weight + "",
                            oprindelse: ""
                            );
                        output.Add(newEntry);
                    }
                }
                catch { }
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