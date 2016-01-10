using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace ExportToExcelV2
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Handles the Click event of the btnExport control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            GenerateReport();
        }

        /// <summary>
        /// Generates the report.
        /// </summary>
        private static void GenerateReport()
        {
            var startTime = DateTime.Now.Second;
            using (ExcelPackage p = new ExcelPackage())
            {

                //set the workbook properties and add a default sheet in it
                SetWorkbookProperties(p);
                //Create a sheet
                ExcelWorksheet ws = CreateSheet(p,"Sample Sheet");
                DataTable dt = CreateDataTable(); //My Function which generates DataTable
                List<TestData> testData = BuildList();
                ws.InsertRow(1, testData.Count);
                ws.InsertColumn(1, 7);
                
                //Merging cells and create a center heading for out table
                ws.Cells[1, 1].Value = "Sample DataTable Export";
                ws.Cells[1, 1, 1, 6].Merge = true;
                ws.Cells[1, 1, 1, 6].Style.Font.Bold = true;
                ws.Cells[1, 1, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int rowIndex = 2;

                CreateHeader(ws, ref rowIndex);
                CreateData(ws, ref rowIndex, testData);
                CreateFooter(ws, ref rowIndex);

                AddComment(ws, 5, 10, "Zeeshan Umar's Comments", "Zeeshan Umar");

                ws.Cells.AutoFitColumns();
                //string path = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Application.StartupPath)), "Zeeshan Umar.jpg");
                //AddImage(ws, 10, 0, path);

                //AddCustomShape(ws, 10, 7, eShapeStyle.Ellipse, "Text inside Ellipse.");

                //Generate A File with Random name
                Byte[] bin = p.GetAsByteArray();
                string file = Guid.NewGuid().ToString() + ".xlsx";
                File.WriteAllBytes(@"C:\temp\" + file, bin);
                var endTime = DateTime.Now.Second;

                MessageBox.Show("All done it took " + (endTime - startTime) + " seconds for 25000 rows");
            }
        }

        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[1];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }

        /// <summary>
        /// Sets the workbook properties and adds a default sheet.
        /// </summary>
        /// <param name="p">The p.</param>
        /// <returns></returns>
        private static void SetWorkbookProperties(ExcelPackage p)
        {
            //Here setting some document properties
            p.Workbook.Properties.Author = "Zeeshan Umar";
            p.Workbook.Properties.Title = "EPPlus Sample";

            
        }

        private static void CreateHeader(ExcelWorksheet ws, ref int rowIndex)
        {
            int colIndex = 1;
            List<string> columns = new List<string>();
            columns.Add("Company");
            columns.Add("Manufacturer");
            columns.Add("Product");
            columns.Add("Rebate Rate");
            columns.Add("Quantity");
            columns.Add("Total Rabate");
            foreach (var dc in columns) //Creating Headings
            {
                var cell = ws.Cells[rowIndex, colIndex];

                //Setting the background color of header cells to Gray
                var fill = cell.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(Color.Gray);

                //Setting Top/left,right/bottom borders.
                var border = cell.Style.Border;
                border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;

                //Setting Value in cell
                cell.Value = dc; 
                colIndex++;
            }
        }

        private static void CreateData(ExcelWorksheet ws, ref int rowIndex, List<TestData> testData)
        {
            int colIndex=0;
            List<string> columns = new List<string>();
            columns.Add("Company");
            columns.Add("Manufacturer");
            columns.Add("Product");
            columns.Add("Rebate Rate");
            columns.Add("Quantity");
            columns.Add("Total Rabate");

            foreach (var td in testData)
            {
                colIndex = 1;
                rowIndex++;
                var cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.Company;
                //Setting borders of cell
                var border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;

                cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.Manufacturer;
                //Setting borders of cell
                border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;

                cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.ProductDescription;
                //Setting borders of cell
                border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;

                cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.RebateAmount;
                cell.Style.Numberformat.Format = "#,##0.00";
                //Setting borders of cell
                border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;

                cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.Quantity;
                cell.Style.Numberformat.Format = "#,##0";
                //Setting borders of cell
                border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;

                cell = ws.Cells[rowIndex, colIndex];
                //Setting Value in cell
                cell.Value = td.TotalRebate;
                cell.Style.Numberformat.Format = "$#,##0.00";
                //Setting borders of cell
                border = cell.Style.Border;
                border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                colIndex++;
            }
        }

        private static void CreateFooter(ExcelWorksheet ws, ref int rowIndex)
        {
            var cell = ws.Cells[rowIndex + 1, 4];
            //Setting Sum Formula
            cell.Formula = "Sum(" + ws.Cells[3, 4].Address + ":" + ws.Cells[rowIndex, 4].Address + ")";
            cell.Style.Numberformat.Format = "$#,##0.00";
            //Setting Background fill color to Gray
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.Gray);

            cell = ws.Cells[rowIndex + 1, 5];
            //Setting Sum Formula
            cell.Formula = "Sum(" + ws.Cells[3, 5].Address + ":" + ws.Cells[rowIndex, 5].Address + ")";
            cell.Style.Numberformat.Format = "#,##0";
            //Setting Background fill color to Gray
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.Gray);

            cell = ws.Cells[rowIndex + 1, 6];
            //Setting Sum Formula
            cell.Formula = "Sum(" + ws.Cells[3, 6].Address + ":" + ws.Cells[rowIndex, 6].Address + ")";
            cell.Style.Numberformat.Format = "$#,##0.00";
            //Setting Background fill color to Gray
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.Gray);
        }

        /// <summary>
        /// Adds the custom shape.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="shapeStyle">Shape style</param>
        /// <param name="text">Text for the shape</param>
        private static void AddCustomShape(ExcelWorksheet ws, int colIndex, int rowIndex, eShapeStyle shapeStyle, string text)
        {
            ExcelShape shape = ws.Drawings.AddShape("cs" + rowIndex.ToString() + colIndex.ToString(), shapeStyle);
            shape.From.Column = colIndex;
            shape.From.Row = rowIndex;
            shape.From.ColumnOff = Pixel2MTU(5);
            shape.SetSize(100, 100);
            shape.RichText.Add(text);
        }

        /// <summary>
        /// Adds the image in excel sheet.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="filePath">The file path</param>
        private static void AddImage(ExcelWorksheet ws, int columnIndex, int rowIndex, string filePath)
        {
            //How to Add a Image using EP Plus
            Bitmap image = new Bitmap(filePath);
            ExcelPicture picture = null;
            if (image != null)
            {
                picture = ws.Drawings.AddPicture("pic" + rowIndex.ToString() + columnIndex.ToString(), image);
                picture.From.Column = columnIndex;
                picture.From.Row = rowIndex;
                picture.From.ColumnOff = Pixel2MTU(2); //Two pixel space for better alignment
                picture.From.RowOff = Pixel2MTU(2);//Two pixel space for better alignment
                picture.SetSize(100, 100);
            }
        }

        /// <summary>
        /// Adds the comment in excel sheet.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="comments">Comment text</param>
        /// <param name="author">Author Name</param>
        private static void AddComment(ExcelWorksheet ws, int colIndex, int rowIndex, string comment, string author)
        {
            //Adding a comment to a Cell
            var commentCell = ws.Cells[rowIndex, colIndex];
            commentCell.AddComment(comment, author);
        }

        /// <summary>
        /// Pixel2s the MTU.
        /// </summary>
        /// <param name="pixels">The pixels.</param>
        /// <returns></returns>
        public static int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }

        /// <summary>
        /// Creates the data table with some dummy data.
        /// </summary>
        /// <returns>DataTable</returns>
        private static DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < 10; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            for (int i = 0; i < 10; i++)
            {
                DataRow dr = dt.NewRow();
                foreach (DataColumn dc in dt.Columns)
                {
                    dr[dc.ToString()] = i;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        private static List<TestData> BuildList()
        {
            var testData = new List<TestData>();

            for (int i = 1; i < 15000; i++)
            {
                testData.Add(new TestData
                {
                    Company = i % 3 == 0 ? "Wilco" : i % 2 == 0 ? "Hess" : "Speedway",
                    ProductDescription = i % 3 == 0 ? "Kisses" : i % 2 == 0 ? "Hershey Almond Bar" : "Hershey Bar",
                    Manufacturer = "Hershey",
                    Quantity = i + i % 2,
                    RebateAmount = Convert.ToDecimal(0.04),
                    TotalRebate = Convert.ToDecimal(200 + (2.12 * i))
                });
            }
            return testData;
        }
        
    }

    public class TestData
    {
        public string Company { get; set; }
        public string Manufacturer { get; set; }
        public string ProductDescription { get; set; }
        public decimal RebateAmount { get; set; }
        public decimal TotalRebate { get; set; }
        public int Quantity { get; set; }
    }
}
