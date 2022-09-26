using CommunityToolkit.HighPerformance;
using Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Diagnostics;
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
        private void acTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void acButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookAC;
            Excel._Worksheet worksheetAC;
            Excel._Worksheet infosheetAC;
            Excel.Range rangeAC;

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
        private void bcTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void bcButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookBC;
            Excel._Worksheet worksheetBC;
            Excel._Worksheet infosheetBC;
            Excel.Range rangeBC;

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
        private void bcpbageriTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void bcpbageriButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookBCP;
            Excel._Worksheet worksheetBCP;
            Excel._Worksheet infosheetBCP;
            Excel.Range rangeBCP;

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
        private void dagrofaTextbox_TextChanged(object sender, EventArgs e)
        {

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
        private void emmerysTextbox_TextChanged(object sender, EventArgs e)
        {

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
        private void frisksnitTextbox_TextChanged(object sender, EventArgs e)
        {

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
        private void grøntgrossistenTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void grøntgrossistenButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookGrøntGrossisten;
            Excel._Worksheet worksheetGrøntGrossisten;
            Excel._Worksheet infosheetGrøntGrossisten;
            Excel.Range rangeGrøntGrossisten;

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
        private void hørkramTextbox_TextChanged(object sender, EventArgs e)
        {

        }
        private void hørkramButton_Click(object sender, System.EventArgs e)
        {

            Excel._Workbook workbookHørkram;
            Excel._Worksheet worksheetHørkram;
            Excel._Worksheet infosheetHørkram;
            Excel.Range rangeHørkram;

            try
            {
                workbookHørkram = excelProgram.Workbooks.Open(@"C:\Users\KOM\Documents\Academy opgaver\Automatisering af hospitalsdata\Data til del 1\Hørkram.xlsx");
                worksheetHørkram = workbookHørkram.Sheets[2];
                infosheetHørkram = workbookHørkram.Sheets[1];
                rangeHørkram = worksheetHørkram.UsedRange;
                int rowCountHørkram = rangeHørkram.Rows.Count;
                int colCountHørkram = rangeHørkram.Columns.Count;

                // Imports the date from the Hørkram worksheet
                DateTime dateHørkram = DateTime.Parse(infosheetHørkram.Cells[5, 2].Text);

                // Imports the cell data from the Hørkram sheet as an array of Objects
                Object[,] arrayHørkram = rangeHørkram.get_Value();

                // Creates a List of String arrays for every row in the Hørkram worksheet.
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
                listHørkram.RemoveRange(0, 2); // Header entries in row 1 and 2
                listHørkram.RemoveAll(s => s[4].Contains("Non food") // Entries for non-food items
                || s[4].Contains("Hjælpevarenumre")
                || s[4].Contains("Engangsmateriale")
                || s[4].Contains("storkøkkentilbehør"));

                rangeMerged = worksheetMerged.get_Range("A2", "M" + listHørkram.Count + 1);

                object[,] arrayMerged = rangeMerged.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);
                
                for (int row = 0; row<listHørkram.Count; row++)
                {
                    arrayMerged[row + 1, 1] = dateHørkram.Year;
                    arrayMerged[row + 1, 2] = (dateHørkram.Month)/ 3 + 1;
                    arrayMerged[row + 1, 3] = listHørkram[row].GetValue(1);
                    arrayMerged[row + 1, 4] = listHørkram[row].GetValue(4);
                    arrayMerged[row + 1, 5] = "Hørkram";
                    arrayMerged[row + 1, 6] = listHørkram[row].GetValue(5);
                    if(listHørkram[row].GetValue(6) as String == "J")
                    {
                        arrayMerged[row + 1, 7] = "Øko";
                    }
                    if (listHørkram[row].GetValue(6) as String == "N")
                    {
                        arrayMerged[row + 1, 7] = "Konv";
                    }
                    arrayMerged[row + 1, 8] = listHørkram[row].GetValue(3);
                    arrayMerged[row + 1, 9] = float.Parse(listHørkram[row].GetValue(10) as String)/ float.Parse(listHørkram[row].GetValue(9) as String);
                    arrayMerged[row + 1, 10] = listHørkram[row].GetValue(10);
                    arrayMerged[row + 1, 11] = listHørkram[row].GetValue(11);
                    arrayMerged[row + 1, 12] = listHørkram[row].GetValue(13);
                    arrayMerged[row + 1, 13] = listHørkram[row].GetValue(7);
                }

                rangeMerged.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, arrayMerged);
                rangeMerged = worksheetMerged.UsedRange;

                //Format the cells.
                worksheetMerged.get_Range("A2", "V" + (listHørkram.Count)).Font.Name = "Calibri";
                worksheetMerged.get_Range("A2", "V" + (listHørkram.Count)).Font.Size = 11;

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
        public void MRCO(Object comObject) // Based on code from breezetree.com/blog/
        {
            if (comObject != null)
            {
                Marshal.ReleaseComObject(comObject);
                comObject = null;
            }
        }

    }
}
