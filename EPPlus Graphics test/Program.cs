using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

using System;
using System.Drawing;
using System.IO;

namespace EPPlus_Graphics_test
{
	class Program
	{
        private const string OutputFolder= @"C:\Temp\EPPlus\";

        private const string ImageFileName = "AD-UNI-L-DFA";

        //default column width 53px and row height 17px
        private const double DefaultColWidth = 7.56;
        private const double DefaultRowHeight = 12.75;

        //font size for entire sheet except title
        private const float FontSize = 10;
        private const string FontName = "Calibri";

        //title font size
        private const int FontSizeTitle = 11;

        //page margins
        private const double TopMargin = .75;
        private const double BottomMargin = .75;
        private const double LeftMargin = .25;
        private const double RightMargin = .25;
        private const double HeaderMargin = .3;
        private const double FooterMargin = .3;

		static readonly string ImageFullFileName = @$"C:\Temp\EPPlus\{ImageFileName}.png";

        private static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var XlOutputFileName = $"{OutputFolder}Test - {DateTime.Now:yyyy-MM-dd HH-mm-ss}.xlsx";

            if (File.Exists(XlOutputFileName))
                File.Delete(XlOutputFileName);

            using (var excelPackage = new ExcelPackage(new FileInfo(XlOutputFileName)))
			{
                var xlWorkSheet = AddTestSheet(excelPackage);

                excelPackage.Save();

                xlWorkSheet.Dispose();
            }

            Console.WriteLine("Press any key");
            Console.ReadKey();
        }

		private static ExcelWorksheet AddTestSheet(ExcelPackage excelPackage)
		{
			var xlWorkSheet = excelPackage.Workbook.Worksheets.Add("Test Images");

            //set column width and row height
            xlWorkSheet.DefaultColWidth = DefaultColWidth;
            xlWorkSheet.DefaultRowHeight = DefaultRowHeight;

            //set default font, font must be set after DefaultColWidth
            xlWorkSheet.Cells.Style.Font.Size = FontSize;
            xlWorkSheet.Cells.Style.Font.Name = FontName;

            //set page margins
            xlWorkSheet.PrinterSettings.TopMargin = (decimal)TopMargin;
            xlWorkSheet.PrinterSettings.BottomMargin = (decimal)BottomMargin;
            xlWorkSheet.PrinterSettings.LeftMargin = (decimal)LeftMargin;
            xlWorkSheet.PrinterSettings.RightMargin = (decimal)RightMargin;
            xlWorkSheet.PrinterSettings.HeaderMargin = (decimal)HeaderMargin;
            xlWorkSheet.PrinterSettings.FooterMargin = (decimal)FooterMargin;

            xlWorkSheet.View.ShowGridLines = false;
            xlWorkSheet.View.PageLayoutView = true;

			const int leftColumn = 1;
            const int rightColumn = 13;

            //add title
            var currentRow = 1;
            //top line
            xlWorkSheet.Cells[currentRow, leftColumn, currentRow, rightColumn].Style.Border.Top.Style = ExcelBorderStyle.Double;
            //set title font
            xlWorkSheet.Cells[currentRow, 1].Style.Font.Size = FontSizeTitle;

            currentRow += 4;

            //add first image
            using (var image1 = new Bitmap(ImageFullFileName))
			{
				using (var excelImage = xlWorkSheet.Drawings.AddPicture(GetImageName(xlWorkSheet, ImageFileName), image1))
				{
                    xlWorkSheet.Cells[currentRow, leftColumn].Value = excelImage.Name;

                    excelImage.From.Column = leftColumn;
                    excelImage.From.Row = currentRow + 1;

                    excelImage.ChangeCellAnchor(eEditAs.TwoCell);
                    excelImage.EditAs = eEditAs.Absolute;
                }
			}
            currentRow += 10;
            Console.WriteLine($"Image1 added, image count = {xlWorkSheet.Drawings.Count}");

            using (var image2 = new Bitmap(ImageFullFileName))
            {
                using (var excelImage = xlWorkSheet.Drawings.AddPicture(GetImageName(xlWorkSheet, ImageFileName), image2))
                {
                    xlWorkSheet.Cells[currentRow, leftColumn].Value = $"{excelImage.Name} - scaled 50%";

                    excelImage.From.Column = leftColumn;
                    excelImage.From.Row = currentRow + 1;

                    excelImage.ChangeCellAnchor(eEditAs.TwoCell);
                    excelImage.EditAs = eEditAs.Absolute;

                    excelImage.SetSize(50);
                }
            }
            currentRow += 10;
            Console.WriteLine($"Image2 added, image count = {xlWorkSheet.Drawings.Count}");

            var nameToRemove = string.Empty;
            using (var image3 = new Bitmap(ImageFullFileName))
            {
                using (var excelImage = xlWorkSheet.Drawings.AddPicture(GetImageName(xlWorkSheet, ImageFileName), image3))
                {
                    xlWorkSheet.Cells[currentRow, leftColumn].Value = excelImage.Name;

                    excelImage.From.Column = leftColumn;
                    excelImage.From.Row = currentRow + 1;
                    nameToRemove = excelImage.Name;

					excelImage.ChangeCellAnchor(eEditAs.TwoCell);
                    excelImage.EditAs = eEditAs.Absolute;
                }
            }
            currentRow += 10;
            Console.WriteLine($"Image3 added, image count = {xlWorkSheet.Drawings.Count}");

            var index = xlWorkSheet.Drawings.Count - 1;

            //remove drawing by name
            try
			{
                xlWorkSheet.Drawings.Remove(nameToRemove);
			}
			catch (Exception ex)
			{
                Console.WriteLine($"Error when trying to remove image by name\nImage name: {nameToRemove}\nError: {ex.Message}\n");
            }
            //remove drawing by name
            try
            {
                xlWorkSheet.Drawings.Remove(xlWorkSheet.Drawings[index].Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error when trying to remove image by name\nImage name: {xlWorkSheet.Drawings[index].Name}\nError: {ex.Message}\n");
            }

            //remove drawing by index
            try
            {
                xlWorkSheet.Drawings.Remove(index);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error when trying to remove image by index\nImage index: {index}\nError: {ex.Message}\n");
            }

            //remove drawing as drawing
            try
            {
                xlWorkSheet.Drawings.Remove(xlWorkSheet.Drawings[index]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error when trying to remove image as drawing index\nxlWorkSheet.Drawings[index].Name: {xlWorkSheet.Drawings[index].Name}\nError: {ex.Message}\n");
            }

            Console.WriteLine($"Image count = {xlWorkSheet.Drawings.Count}");

            //adding one more image
            using (var image4 = new Bitmap(ImageFullFileName))
            {
                using (var excelImage = xlWorkSheet.Drawings.AddPicture(GetImageName(xlWorkSheet, ImageFileName), image4))
                {
                    xlWorkSheet.Cells[currentRow, leftColumn].Value = excelImage.Name;

                    excelImage.From.Column = leftColumn;
                    excelImage.From.Row = currentRow + 1;

                    excelImage.ChangeCellAnchor(eEditAs.TwoCell);
                    excelImage.EditAs = eEditAs.Absolute;
                }
            }
            Console.WriteLine($"Image4 added, image count = {xlWorkSheet.Drawings.Count}");

            return xlWorkSheet;
		}

        private static string GetImageName(ExcelWorksheet sheet, string fileName)
        {
            var imageNumber = sheet.Drawings.Count;
            return $"image{imageNumber}_{fileName}";
        }
    }
}
