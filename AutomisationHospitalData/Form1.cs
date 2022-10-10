using CommunityToolkit.HighPerformance;
using Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomisationHospitalData
{
    public partial class Form1 : Form
    {
        Excel.Application excelProgram;
        Excel._Workbook workbookMerged;
        Excel._Worksheet worksheetMerged;
        Excel.Range rangeMerged;

        // paths for companies supplying folders of excel sheets
        List<String> pathAC = new List<String>();
        List<String> pathDagrofa = new List<String>();
        List<String> pathFrisksnit = new List<String>();

        // paths for companies supplying individual excel sheets
        string pathBC = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\BC.xlsx";
        string pathCBP = @"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\CBP bageri.xlsx";
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
        private void createNewExcelTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void createNewExcelButton_Click(object sender, System.EventArgs e)
        {
        }
        private void buttonACPath_Click(object sender, EventArgs e)
        {
            this.openACPathDialog.Multiselect = true;
            this.openACPathDialog.Title = "Select AC files";

            if (openACPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathAC = openACPathDialog.FileNames.ToList();
                buttonACPath.Text = openACPathDialog.FileName;
            }
        }
        private void acButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookAC;
            Excel._Worksheet worksheetAC;
            Excel.Range rangeAC;

            try
            {
                foreach(String fileAC in pathAC)
                {
                    workbookAC = excelProgram.Workbooks.Open(fileAC);
                    worksheetAC = workbookAC.Sheets[1];
                    rangeAC = worksheetAC.UsedRange;

                    Object[,] firstColumn = rangeAC.get_Value();

                    int rowCountAC = 2;
                    for(int count = 2; count < rangeAC.Rows.Count; count++)
                    {
                        try // 
                        {
                            if (float.Parse(firstColumn[count,1].ToString()) > 0)
                            {
                                rowCountAC++;
                            }
                        }
                        catch (NullReferenceException) // "Catch" in case the cell's value is Null
                        {
                            break;
                        }
                    }
                    rangeAC = worksheetAC.get_Range("A1", "H"+ rowCountAC);

                    int colCountAC = rangeAC.Columns.Count;

                    int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                    // Imports the cell data from the Hørkram sheet as an array of Objects
                    Object[,] arrayAC = rangeAC.get_Value();

                    // Creates a List of String arrays for every rowOld in the AC worksheet.
                    // Amount of rows as a List to allow for deletion of irrelevant entries.
                    List<String[]> listAC = new List<String[]>();

                    // For every row in the imported AC Object array, copy its value to the corresponding String in the List of String arrays
                    for (int row = 0; row < rowCountAC; row++)
                    {
                        Debug.WriteLine(row);
                        listAC.Add(new string[14]);
                        for (int col = 0; col < 8; col++)
                        {
                            try // "Try" because the cell's value can be Null
                            {
                                listAC[row].SetValue(arrayAC[row + 1, col + 1].ToString(), col);
                            }
                            catch (NullReferenceException) // "Catch" in case the cell's value is Null
                            {
                                listAC[row].SetValue("", col);
                            }
                        }
                    }

                    // Deletion of irrelevant entries from the List of String arrays
                    listAC.RemoveRange(0, 1); // Header entries in row 1

                    rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (usedRowsMerged + listAC.Count - 1));

                    object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                    // Sets the values in the AC Object Array
                    for (int row = 0; row < listAC.Count-1; row++)
                    {
                        Debug.WriteLine(row);
                        arrayMerged[row + 1, 1] = "Ikke oplyst"; // År
                        arrayMerged[row + 1, 2] = "Ikke oplyst"; // Kvartal
                        arrayMerged[row + 1, 3] = listAC[row].GetValue(1); // Hospital
                        arrayMerged[row + 1, 4] = ""; // Råvarekategori
                        arrayMerged[row + 1, 5] = "AC"; // Leverandør
                        arrayMerged[row + 1, 6] = ""; // Råvare

                        String[] nameSplitAC1 = (listAC[row].GetValue(3) as String).Split(' ');
                        String[] nameSplitAC2 = (listAC[row].GetValue(3) as String).Split('(');
                        
                        if (nameSplitAC1.First() == "ØKO")
                        {
                            arrayMerged[row + 1, 7] = "Øko"; // øko
                        }
                        else
                        {
                            arrayMerged[row + 1, 7] = "Konv"; // konv
                        }
                        arrayMerged[row + 1, 8] = listAC[row].GetValue(3); // Varianter/opr
                        arrayMerged[row + 1, 9] = listAC[row].GetValue(7); // Pris pr enhed
                        arrayMerged[row + 1, 10] = listAC[row].GetValue(6); // Pris i alt
                        arrayMerged[row + 1, 11] = listAC[row].GetValue(5); // Kg

                        float floatTotalPrice = float.Parse(listAC[row].GetValue(6).ToString());
                        float floatWeight = float.Parse(listAC[row].GetValue(5).ToString());

                        arrayMerged[row + 1, 12] = floatTotalPrice/floatWeight + ""; // Kilopris
                        if (nameSplitAC2.Last().Length > 2)
                        {
                            arrayMerged[row + 1, 13] = nameSplitAC2.Last().Substring(0,3); // Oprindelse
                        }
                        else
                        {
                            arrayMerged[row + 1, 13] = "Ikke oplyst"; // Oprindelse
                        }
                    }

                    rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                    rangeMerged = worksheetMerged.UsedRange;

                    //Format the cells.
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listAC.Count - 1)).Font.Name = "Calibri";
                    worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (usedRowsMerged + listAC.Count - 1)).Font.Size = 11;

                    //AutoFit columns A:V.
                    rangeMerged = worksheetMerged.get_Range("A1", "M1");
                    rangeMerged.EntireColumn.AutoFit();


                    //Make sure Excel is visible and give the user control
                    //of Microsoft Excel's lifetime.
                    excelProgram.Visible = true;
                    excelProgram.UserControl = true;

                    // Releasing the Excel interop objects
                    workbookAC.Close(false);
                    MRCO(workbookAC);
                    MRCO(worksheetAC);
                    MRCO(rangeAC);
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
        private void buttonBCPath_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Title = "Select BC File";
            if (openBCPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathBC = openBCPathDialog.FileName;
                buttonBCPath.Text = openBCPathDialog.FileName;
            }
        }
        private void bcButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookBC;
            Excel._Worksheet worksheetBC;
            Excel._Worksheet infosheetBC;
            Excel.Range rangeBC;

            try
            {
                workbookBC = excelProgram.Workbooks.Open(pathBC);
                worksheetBC = workbookBC.Sheets[1];
                infosheetBC = workbookBC.Sheets[2];
                rangeBC = worksheetBC.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountBC = rangeBC.Rows.Count;
                int colCountBC = rangeBC.Columns.Count;

                // Imports the date from the BC worksheet
                String infostringBC = worksheetBC.Cells[2, 2].Text;
                String[] dateBCString = infostringBC.Split(new string[] { ".." }, StringSplitOptions.None);

                DateTime dateBC = DateTime.Parse(dateBCString[1]);

                // Imports the cell data from the Hørkram sheet as an array of Objects
                Object[,] arrayBC= rangeBC.get_Value();

                // Creates a List of String arrays for every rowOld in the BC worksheet.
                // Amount of rows as a List to allow for deletion of irrelevant entries.
                List<List<String>> listBC = new List<List<String>>();

                int rowOld = 5;
                int rowNew = 0;

                while (rowOld < rowCountBC - 7)
                {
                    Boolean hospitalEnd = false;
                    Boolean categoryEnd = false;
                    Boolean skipping = false;

                    string currentHospital = arrayBC[rowOld + 1,2].ToString();
                    rowOld++;
                    rowOld++;
                    string currentCategory = arrayBC[rowOld + 1, 2].ToString();
                    rowOld++;
                    rowOld++;

                    while (!hospitalEnd)
                    {
                        while (!categoryEnd)
                        {
                            if (float.Parse(arrayBC[rowOld + 1, 7].ToString()) > 0)
                            //if (!(arrayBC[rowOld + 1, 6] == null))
                            {
                                listBC.Add(new List<String>());
                                listBC[rowNew].Add("" + dateBC.Year); // 0 år
                                listBC[rowNew].Add("" + (dateBC.Month / 3)); // 1 Kvartal
                                listBC[rowNew].Add(currentHospital); // 2 Hospital
                                listBC[rowNew].Add(currentCategory); // 3 Råvarekategori
                                listBC[rowNew].Add("BC"); // 4 Leverandør
                                try 
                                {
                                    listBC[rowNew].Add(arrayBC[rowOld + 1, 19].ToString()); // 5 Råvare
                                }
                                catch
                                {
                                    listBC[rowNew].Add("");
                                }
                                try // "Try" to check whether this cell is empty
                                {
                                    arrayBC[rowOld + 1, 9].ToString();
                                    listBC[rowNew].Add("Øko"); // 6 Øko
                                }
                                catch (NullReferenceException) // "Catch" in case the cell is empty
                                {
                                    listBC[rowNew].Add("Konv."); // 6 Konv.
                                }
                                listBC[rowNew].Add(arrayBC[rowOld + 1, 2].ToString()); // 7 Varianter/por
                                listBC[rowNew].Add(arrayBC[rowOld + 1, 8].ToString()); // 8 Pris pr enhed
                                listBC[rowNew].Add(arrayBC[rowOld + 1, 7].ToString()); // 9 Pris i alt
                                try // "Try" because this cell is empty if non-ecological
                                {
                                    listBC[rowNew].Add(arrayBC[rowOld + 1, 9].ToString()); // 10 Kilo
                                }
                                catch (NullReferenceException) // "Catch" in case the cell is empty
                                {
                                    try // "Try" because this cell is empty if it's "withheld"
                                    {
                                        listBC[rowNew].Add(arrayBC[rowOld + 1, 10].ToString()); // 10 Kilo
                                    }
                                    catch (NullReferenceException) // "Catch" in case the cell is empty
                                    {
                                        listBC[rowNew].Add(arrayBC[rowOld + 1, 11].ToString()); // 10 Kilo
                                    }
                                }
                                float kiloBC = float.Parse(listBC[rowNew][10]);
                                float priceBC = float.Parse(listBC[rowNew][9]);
                                float kilopriceBC = priceBC / kiloBC;
                                listBC[rowNew].Add("" + kilopriceBC); // 11 Kilopris
                                try // "Try" to check whether this cell is empty
                                {
                                    listBC[rowNew].Add(arrayBC[rowOld + 1, 17].ToString()); // 12 Oprindelse
                                }
                                catch (NullReferenceException) // "Catch" in case the cell is empty
                                {
                                    try // "Try" to check whether this cell is empty
                                    {
                                        listBC[rowNew].Add(arrayBC[rowOld + 1, 18].ToString()); // 12 Oprindelse
                                    }
                                    catch (NullReferenceException) // "Catch" in case the cell is empty
                                    {
                                        listBC[rowNew].Add("Not supplied"); // 12 Oprindelse
                                    }
                                }
                                rowNew++;
                            }
                            rowOld++;

                            try // "Try" to check if this cell is empty
                            {
                                arrayBC[rowOld + 1, 2].ToString(); 
                            }
                            catch (NullReferenceException) // "Catch" in case the cell is empty
                            {
                                categoryEnd = true;
                            }
                        }
                        rowOld++;

                        try {
                            currentCategory = arrayBC[rowOld + 1, 2].ToString();
                            int categoryNumber = int.Parse(currentCategory.Split(' ')[0]);
                            categoryEnd = false;

                            if (categoryNumber > 88)
                            {
                                hospitalEnd = true;
                            }
                            else
                            {
                                rowOld++;
                                rowOld++;
                            }
                        }
                        catch
                        {
                            hospitalEnd = true;
                        }
                    }
                    skipping = true;
                    while(skipping)
                    {
                        rowOld++;
                        try 
                        {
                            if (arrayBC[rowOld+1,2].ToString().Contains("Total beløb"))
                            {
                                skipping = false;
                                rowOld++;
                                rowOld++;
                                rowOld++;
                            }
                        }
                        catch (NullReferenceException) // "Catch" in case the cell is empty
                        {
                        }
                    }
                }

                rangeMerged = worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "M" + (listBC.Count + usedRowsMerged));
                object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                for (int row = 0; row < listBC.Count; row++)
                {
                    for (int column = 0; column < 13; column++)
                    {
                        arrayMerged[row + 1, column + 1] = listBC[row][column];
                    }
                }

                rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (listBC.Count + usedRowsMerged)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A" + (usedRowsMerged + 1), "V" + (listBC.Count + usedRowsMerged)).Font.Size = 11;

                //AutoFit columns A:V.
                rangeMerged = worksheetMerged.get_Range("A1", "M1");
                rangeMerged.EntireColumn.AutoFit();


                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                excelProgram.Visible = true;
                excelProgram.UserControl = true;

                // Releasing the Excel interop objects
                workbookBC.Close(false);
                MRCO(workbookBC);
                MRCO(worksheetBC);
                MRCO(infosheetBC);
                MRCO(rangeBC);

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
        private void cbpbageriButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookCBP;
            Excel._Worksheet worksheetCBP;
            Excel._Worksheet infosheetCBP;
            Excel.Range rangeCBP;

            try
            {

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
        private void dagrofaButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookDagrofa;
            Excel._Worksheet worksheetDagrofa;
            Excel._Worksheet infosheetDagrofa;
            Excel.Range rangeDagrofa;

            try
            {

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
        private void emmerysButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookEmmerys;
            Excel._Worksheet worksheetEmmerys;
            Excel._Worksheet infosheetEmmerys;
            Excel.Range rangeEmmerys;

            try
            {

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
        private void frisksnitButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookFrisksnit;
            Excel._Worksheet worksheetFrisksnit;
            Excel._Worksheet infosheetFrisksnit;
            Excel.Range rangeFrisksnit;

            try
            {

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
        private void grøntgrossistenButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookGrøntGrossisten;
            Excel._Worksheet worksheetGrøntGrossisten;
            Excel._Worksheet infosheetGrøntGrossisten;
            Excel.Range rangeGrøntGrossisten;

            try
            {
                workbookGrøntGrossisten = excelProgram.Workbooks.Open(pathGrøntGrossisten);
                worksheetGrøntGrossisten = workbookGrøntGrossisten.Sheets[2];
                infosheetGrøntGrossisten = workbookGrøntGrossisten.Sheets[1];
                rangeGrøntGrossisten = worksheetGrøntGrossisten.UsedRange;

                int usedRowsMerged = worksheetMerged.UsedRange.Rows.Count;

                int rowCountHørkram = rangeGrøntGrossisten.Rows.Count;
                int colCountHørkram = rangeGrøntGrossisten.Columns.Count;
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
        private void buttonHørkramPath_Click(object sender, EventArgs e)
        {
            this.openHørkramPathDialog.Title = "Select Hørkram file";
            if (openHørkramPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathHørkram = openHørkramPathDialog.FileName;
                buttonHørkramPath.Text = openHørkramPathDialog.FileName;
            }
        }
        private void hørkramButton_Click(object sender, System.EventArgs e)
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
                for (int row = 0; row<listHørkram.Count; row++)
                {
                    arrayMerged[row + 1, 1] = dateHørkram.Year; // År
                    arrayMerged[row + 1, 2] = (dateHørkram.Month)/ 3; // Kvartal
                    arrayMerged[row + 1, 3] = listHørkram[row].GetValue(1); // Hospital
                    arrayMerged[row + 1, 4] = listHørkram[row].GetValue(4); // Råvarekategori
                    arrayMerged[row + 1, 5] = "Hørkram"; // Leverandør
                    arrayMerged[row + 1, 6] = listHørkram[row].GetValue(5); // Råvare
                    if(listHørkram[row].GetValue(6) as String == "J") // konv/øko
                    {
                        arrayMerged[row + 1, 7] = "Øko";
                    }
                    if (listHørkram[row].GetValue(6) as String == "N")
                    {
                        arrayMerged[row + 1, 7] = "Konv";
                    }
                    arrayMerged[row + 1, 8] = listHørkram[row].GetValue(3); // Varianter/opr
                    arrayMerged[row + 1, 9] = float.Parse(listHørkram[row].GetValue(10) as String)/ float.Parse(listHørkram[row].GetValue(9) as String); // Pris pr enhed
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
        public void MRCO(Object comObject) // Based on code from breezetree.com/blog/
        {
            if (comObject != null)
            {
                Marshal.ReleaseComObject(comObject);
                comObject = null;
            }
        }

        private OpenFileDialog openACPathDialog = new OpenFileDialog();
        private OpenFileDialog openBCPathDialog = new OpenFileDialog();
        private OpenFileDialog openGrøntGrossistenPathDialog = new OpenFileDialog();
        private OpenFileDialog openHørkramPathDialog = new OpenFileDialog();

        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            MRCO(excelProgram);
            MRCO(workbookMerged);
            MRCO(worksheetMerged);
            MRCO(rangeMerged);
        }

        private void buttonCBPBageriPath_Click(object sender, EventArgs e)
        {

        }

        private void buttonDagrofaPath_Click(object sender, EventArgs e)
        {

        }

        private void buttonEmmerysPath_Click(object sender, EventArgs e)
        {

        }

        private void buttonFriskSnitPath_Click(object sender, EventArgs e)
        {

        }

        private void buttonGrøntGrossistenPath_Click(object sender, EventArgs e)
        {
            this.openGrøntGrossistenPathDialog.Title = "Select Grønt Grossisten file";
            if (openGrøntGrossistenPathDialog.ShowDialog() == DialogResult.OK)
            {
                pathGrøntGrossisten = openGrøntGrossistenPathDialog.FileName;
                buttonGrøntGrossistenPath.Text = openGrøntGrossistenPathDialog.FileName;
            }
        }
    }
}
