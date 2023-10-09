// See https://aka.ms/new-console-template for more information
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.Extensions.Configuration;

var configuration = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

using (var service = new SheetsService(new BaseClientService.Initializer()
    {
        ApplicationName = "WallingfordHoops",
        ApiKey = configuration["GoogleSheetsApiKey"],
    }))
{
    // Begin timing
    var stopwatch = new System.Diagnostics.Stopwatch();
    stopwatch.Start();

    var getSpreadsheetRequest = service
        .Spreadsheets
        .Get("1IpB80F49aSd8C5m1VvNJm2T12njdVSjzWl92k0cqpPM");
    var spreadsheet = getSpreadsheetRequest.Execute();

    var rangesForSheetsWithGameData = new List<string>();

    Console.WriteLine("All sheets in spreadsheet:");
    foreach (var sheet in spreadsheet.Sheets)
    {
        Console.WriteLine(sheet.Properties.Title);
        if (sheet.Properties.Title.StartsWith("2023"))
        {
            rangesForSheetsWithGameData.Add($"{sheet.Properties.Title}!A1:O17");
        }
    }

    Console.WriteLine();

    getSpreadsheetRequest.IncludeGridData = true;
    getSpreadsheetRequest.Ranges = rangesForSheetsWithGameData;
    var sheetsWithGameData = getSpreadsheetRequest
        .Execute()
        .Sheets;

    Console.WriteLine("All sheets with game data:");
    foreach (var sheet in sheetsWithGameData)
    {
        Console.WriteLine(sheet.Properties.Title);
        foreach (var gridData in sheet.Data)
        {
            foreach (var row in gridData.RowData)
            {
                var rowValues = row.Values.Select(cellData => cellData.FormattedValue);
                Console.WriteLine(String.Join(", ", rowValues));
            }
        }
        break; // only process one sheet for now
    }

    Console.WriteLine();

    // Stop timing
    stopwatch.Stop();
    Console.WriteLine("Time taken : {0}", stopwatch.Elapsed);
}