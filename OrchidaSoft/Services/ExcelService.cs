namespace OrchidaSoft.Services
{
    using OfficeOpenXml;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using OrchidaSoft.Models;
    using Microsoft.AspNetCore.Mvc;
    using System;

    public class ExcelService
    {
        //public List<TaxData> ImportExcel(Stream fileStream)
        //{
        //    var taxDataList = new List<TaxData>();

        //    using (var package = new ExcelPackage(fileStream))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
        //        var rowCount = worksheet.Dimension.Rows;

        //        for (int row = 2; row <= rowCount; row++) // Assuming the first row is the header
        //        {
        //            var taxData = new TaxData
        //            {
        //                Description = worksheet.Cells[row, 1].Text,
        //                ValueAfterTaxing = decimal.Parse(worksheet.Cells[row, 2].Text)
        //            };
        //            taxDataList.Add(taxData);
        //        }
        //    }

        //    return taxDataList;
        //}

        //public MemoryStream ModifyExcel(List<TaxData> taxDataList)
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Taxes");
        //        worksheet.Cells[1, 1].Value = "Description";
        //        worksheet.Cells[1, 2].Value = "Value After Tax";
        //        worksheet.Cells[1, 3].Value = "Total Value Before Taxing"; // New column

        //        decimal totalValueAfterTaxing = 0;

        //        for (int i = 0; i < taxDataList.Count; i++)
        //        {
        //            worksheet.Cells[i + 2, 1].Value = taxDataList[i].Description;
        //            worksheet.Cells[i + 2, 2].Value = taxDataList[i].ValueAfterTaxing;
        //            worksheet.Cells[i + 2, 3].Value = taxDataList[i].ValueAfterTaxing * 0.8m; // Example calculation
        //            totalValueAfterTaxing += taxDataList[i].ValueAfterTaxing;
        //        }

        //        // Add the total row
        //        worksheet.Cells[taxDataList.Count + 2, 1].Value = "Total";
        //        worksheet.Cells[taxDataList.Count + 2, 2].Value = totalValueAfterTaxing;

        //        var stream = new MemoryStream();
        //        package.SaveAs(stream);
        //        stream.Position = 0; // Reset the stream position
        //        return stream;
        //    }
        //}

        public async Task<MemoryStream> ProcessExcelAsync(Stream fileStream)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(fileStream);
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;

            // Adding header for Total Value Before Taxing and other columns if necessary
            worksheet.Cells[1, 3].Value = "Total Value Before Taxing"; // Total Value Before Taxing header
            worksheet.Cells[1, 4].Value = "Column 4"; // Placeholder for Column 4 header
            worksheet.Cells[1, 5].Value = "Column 5"; // Placeholder for Column 5 header

            decimal totalValueAfterTaxing = 0;

            // Process each row
            for (int row = 2; row <= rowCount; row++)
            {
                var valueAfterTaxing = worksheet.Cells[row, 2].GetValue<decimal>();
                if (valueAfterTaxing == 0) continue; // Skip if the value is 0 or you could have other conditions to skip

                // Calculate Total Value Before Taxing
                decimal totalValueBeforeTaxing = valueAfterTaxing / 1.2m; // Assuming a 20% tax rate
                worksheet.Cells[row, 3].Value = totalValueBeforeTaxing;

                // Example calculations for Column 4 and Column 5 (replace with your logic)
                worksheet.Cells[row, 4].Value = worksheet.Cells[row, 1].Text.Length; // Example: Length of Description
                worksheet.Cells[row, 5].Value = totalValueBeforeTaxing * 10; // Example calculation for Column 5

                // Sum for Total Value After Taxing
                totalValueAfterTaxing += valueAfterTaxing;
            }

            // Add the total row
            int totalRow = rowCount + 1;
            worksheet.Cells[totalRow, 1].Value = "Total";
            worksheet.Cells[totalRow, 2].Value = totalValueAfterTaxing;
            worksheet.Cells[totalRow, 3].Value = ""; // Leave Total Value Before Taxing cell empty or calculate it if necessary
            worksheet.Cells[totalRow, 4].Value = ""; // Leave empty or sum for Column 4
            worksheet.Cells[totalRow, 5].Value = ""; // Leave empty or sum for Column 5

            var resultStream = new MemoryStream();
            await package.SaveAsAsync(resultStream);
            resultStream.Position = 0;

            return resultStream; // Return modified Excel stream
        }

    }
}
