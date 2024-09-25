using System.IO.Packaging;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

class Program
{
    static void Main(string[] args)
    {
        ParseJSON();

        // Path to the Excel file
        string filePath = @"C:\R\September_2024\PGx_Guidelines_Combined.xlsx";

        // Load the Excel file
        using (var workbook = new XLWorkbook(filePath))
        {
            // Get the first worksheet
            var worksheet = workbook.Worksheet(1);

            var drugColumnIndex = -1;
            // Find the column with the header "Drug"
            foreach (var cell in worksheet.Row(1).CellsUsed())
            {
                if (cell.Value.ToString().Trim().Equals("Drug", StringComparison.OrdinalIgnoreCase))
                {
                    drugColumnIndex = cell.Address.ColumnNumber;
                    break;
                }
            }

            // Check if the column was found
            if (drugColumnIndex == -1)
            {
                Console.WriteLine("Column with header 'Drug' not found.");
                return;
            }

            var drugsInfo = new List<object>();
            var drugsResultParsed = new List<object>();

            // Print the values from the "Drug" column and add them to the list
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var cellValue = row.Cell(drugColumnIndex).Value;
                drugsInfo.Add(cellValue);
                Console.WriteLine(cellValue);

                drugsResultParsed.Add(cellValue);
            }           

            var columnsToAdd = 0;

            foreach (var drugList in drugsResultParsed)
            {
                var drugStr = drugList.ToString();
                string[] elements = drugStr.Split(',');

                int count = elements.Length;

                columnsToAdd = count > columnsToAdd ? count : columnsToAdd;
            }

            AddColumnsToExcel(filePath, drugColumnIndex + 1, columnsToAdd);

            var colIndex = 2;

            foreach(var drugList in drugsResultParsed)
            {
                AddValuesToExcel(filePath, drugColumnIndex + 1, colIndex, drugList.ToString());
                colIndex++;
            }

        }

        Console.WriteLine("Finished printing 'Drug' column.");
    }

    public static void ParseJSON()
    {
        // Define the folder with your JSON files
        string inputDirectory = @"C:\R\guidelineAnnotations.json";
        string[] jsonFiles = Directory.GetFiles(inputDirectory, "*.json");

        // Create a list to store the extracted data
        var allGuidelines = new List<GuidelineData>();

        // Loop over each file
        int i = 0;
        foreach (var file in jsonFiles)
        {
            i++;
            Console.WriteLine(i);

            // Read the JSON file
            var jsonData = File.ReadAllText(file);
            var parsedData = JsonConvert.DeserializeObject<RootObject>(jsonData);

            // Extract relevant information
            var guideline = parsedData.Guideline;

            var guidelineData = new GuidelineData
            {
                GuidelineID = guideline.Id,
                Name = guideline.Name,
                Source = guideline.Source,
                GeneId = string.Join(", ", guideline.RelatedGenes.Select(g => g.Id)),
                GeneSymbol = string.Join(", ", guideline.RelatedGenes.Select(g => g.Symbol)),
                Gene = string.Join(", ", guideline.RelatedGenes.Select(g => g.Name)),
                Drug = string.Join(", ", guideline.RelatedChemicals.Select(c => $"{c.Id}, {c.Name}")),
                Pediatric = guideline.Pediatric.ToString(),
                HistoryDate = string.Join(", ", guideline.History.Select(h => h.Date)),
                AlternateDrug = guideline.AlternateDrugAvailable.ToString(),
                OtherGuidance = guideline.OtherPrescribingGuidance.ToString(),
                Recommendations = guideline.TextMarkdown.Html
            };

            if (!string.IsNullOrEmpty(guidelineData.HistoryDate))
            {
                guidelineData.HistoryDate = ProcessAndSortDates(guidelineData.HistoryDate);
            }

            allGuidelines.Add(guidelineData);
        }

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("PGx Guidelines");
            worksheet.Cell(1, 1).InsertTable(allGuidelines); // Requires `allGuidelines` to be in a suitable format
            workbook.SaveAs(@"C:\R\September_2024\PGx_Guidelines_Combined.xlsx");
        }
        

    }

    public static string ProcessAndSortDates(string dates)
    {
        // Split, extract the date part, sort descending, and join into one string
        string sortedDates = string.Join(", ", dates.Split(',')
            .Select(date => date.Trim().Substring(0, 10)) // Take only the year-month-day part
            .OrderByDescending(date => DateTime.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture))); // Sort descending

        return sortedDates;
    }
   
    public static void AddValuesToExcel(string filePath, int startColumn, int startRow, string commaSeparatedValues)
    {
        // Split the string of values by comma
        string[] values = commaSeparatedValues.Split(',');

        // Load the Excel workbook
        using (var workbook = new XLWorkbook(filePath))
        {
            // Get the first worksheet
            var worksheet = workbook.Worksheet(1);

            // Loop through each value and insert it into the Excel sheet
            for (int i = 0; i < values.Length; i++)
            {
                // Calculate the column position (starting from startColumn)
                int currentColumn = startColumn + i;

                // Place the value in the specified row and column
                worksheet.Cell(startRow, currentColumn).Value = values[i].Trim();
            }

            // Save the changes to the file
            workbook.Save();

        }
    }

    public static void AddColumnsToExcel(string filePath, int startColumn, int numberOfColumns)
    {
        // Load the Excel workbook from the file path
        using (var workbook = new XLWorkbook(filePath))
        {
            // Get the first worksheet
            var worksheet = workbook.Worksheet(1);

            // Calculate the number of pairs of 'd' and 'n' columns
            int pairCount = numberOfColumns / 2;

            // Loop through and insert columns and their headers
            for (int i = 0; i < pairCount; i++)
            {
                // Insert new column for "d" and set the header
                worksheet.Column(startColumn + (i * 2)).InsertColumnsBefore(1);
                worksheet.Cell(1, startColumn + (i * 2)).Value = $"DrugId_{i + 1}";

                // Insert new column for "n" and set the header
                worksheet.Column(startColumn + (i * 2) + 1).InsertColumnsBefore(1);
                worksheet.Cell(1, startColumn + (i * 2) + 1).Value = $"DrugName_{i + 1}";
            }

            // Save the modified workbook
            workbook.Save();
        }
    }

}

public class RootObject
{
    public Guideline Guideline { get; set; }
}

public class Guideline
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string Source { get; set; }
    public List<Gene> RelatedGenes { get; set; }
    public List<Chemical> RelatedChemicals { get; set; }
    public bool Pediatric { get; set; }
    public List<History> History { get; set; }
    public bool AlternateDrugAvailable { get; set; }
    public bool OtherPrescribingGuidance { get; set; }
    public TextMarkdown TextMarkdown { get; set; }
}

public class Gene
{
    public string Id { get; set; }
    public string Symbol { get; set; }
    public string Name { get; set; }
}

public class Chemical
{
    public string Id { get; set; }
    public string Name { get; set; }
}

public class History
{
    public string Date { get; set; }
}

public class TextMarkdown
{
    public string Html { get; set; }
}

public class GuidelineData
{
    public string GuidelineID { get; set; }
    public string Name { get; set; }
    public string Source { get; set; }
    public string GeneId { get; set; }
    public string Gene { get; set; }
    public string GeneSymbol { get; set; }
    public string Drug { get; set; }
    public string Pediatric { get; set; }
    public string HistoryDate { get; set; }
    public string AlternateDrug { get; set; }
    public string OtherGuidance { get; set; }
    public string Recommendations { get; set; }
}

