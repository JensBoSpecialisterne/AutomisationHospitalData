using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using CommunityToolkit;
using CommunityToolkit.HighPerformance;
using System.Diagnostics;

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
        private void button1_Click(object sender, System.EventArgs e)
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

                // Creates an array of object lists for every column in the Hørkram worksheet
                // Amount of columns is an array it should remain constant
                // Length of columns is a list to allows for deletion of irrelevant entries
                List<Object>[] arrayHørkram = new List<Object>[14];

                // Iterates over every column in the Hørkram worksheet
                for (int i = 0; i < 14; i++)
                {
                    // Initialises the list for that column's values
                    arrayHørkram[i] = new List<Object>();

                    //Adds every cell in that column to the list
                    foreach(Object cell in (rangeHørkram.Value as object[,]).GetColumn(i))
                        {
                            arrayHørkram[i].Add(cell);
                        }
                }

                // Deletion of irrelevant entries


                // Copies the "Resource" value from the Hørkram worksheet to the merged worksheet
                //rangeMerged = worksheetMerged.get_Range("F2 : F" + (rowCountHørkram - 1));
                //rangeMerged.Value2 = arrayHørkram[5].ToArray();


                /*
               //Sets the "Leverandør" value to Hørkram
               worksheetMerged.get_Range("E2", "E" + (rowCountHørkram - 1)).Value2 = "Hørkram";
               //Sets the "År" value to the year
               worksheetMerged.get_Range("A2", "A" + (rowCountHørkram - 1)).Value2 = DateTime.Parse(infosheetHørkram.Cells[5, 2].Text).Year;

               //Sets the "Kvartal" value to the quarter
               worksheetMerged.get_Range("B2", "B" + (rowCountHørkram - 1)).Value2 = (DateTime.Parse(infosheetHørkram.Cells[5, 2].Text).Month)/3+1;

               // Copies the "Hospital" value from the Hørkram worksheet to the merged worksheet
               Excel.Range hospitalsHørkram = worksheetHørkram.get_Range("B3 : B" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("C2 : C" + (rowCountHørkram - 1));
               hospitalsHørkram.Copy(rangeMerged);

               // Imports the "Øko." value from the Hørkram worksheet as an array
               Excel.Range ecologyHørkram = worksheetHørkram.get_Range("G3 : G" + (rowCountHørkram));

               object[,] ecologyHørkramData = ecologyHørkram.Value as object[,];

               // Rephrases the ecology array to fit the terminology of the merged file
               for (int i = 1; i < ecologyHørkramData.Length+1; i++)
               {
                   if (ecologyHørkramData[i,1] as String== "J")
                   {
                       ecologyHørkramData[i, 1] = "Øko";
                   }
                   if (ecologyHørkramData[i, 1] as String == "N")
                   {
                       ecologyHørkramData[i, 1] = "Konv";
                   }
               }

               // Inserts the ecology array into the merged file
               rangeMerged = worksheetMerged.get_Range("G2 : G" + (rowCountHørkram - 1));
               rangeMerged.Value = ecologyHørkramData;

               // Copies the "Resource category" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("E3 : E" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("D2 : D" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               // Copies the "Resource" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("F3 : F" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("F2 : F" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               // Copies the "Product" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("D3 : D" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("H2 : H" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               // Copies the "Country of origin" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("H3 : H" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("M2 : M" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               // Copies the "Total weight" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("M3 : M" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("K2 : K" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               // Copies the "Price pr. weight" value from the Hørkram worksheet to the merged worksheet
               rangeHørkram = worksheetHørkram.get_Range("N3 : N" + (rowCountHørkram));
               rangeMerged = worksheetMerged.get_Range("L2 : L" + (rowCountHørkram - 1));
               rangeHørkram.Copy(rangeMerged);

               //Format the cells .
               worksheetMerged.get_Range("A2", "V" + (rowCountHørkram - 1)).Font.Name = "Calibri";
               worksheetMerged.get_Range("A2", "V" + (rowCountHørkram - 1)).Font.Size = 11;

               //AutoFit columns A:V.
               rangeMerged = worksheetMerged.get_Range("A1", "D1");
               rangeMerged.EntireColumn.AutoFit();


               //Make sure Excel is visible and give the user control
               //of Microsoft Excel's lifetime.
               excelProgram.Visible = true;
               excelProgram.UserControl = true; */
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

    }
}
