// See https://aka.ms/new-console-template for more information
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using WallingfordHoops;

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
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];
        Console.WriteLine(gameDate);

        PrintGames(gameDate, gridData);

        // PrintRowDataAndGameRowData(gridData, gameDate);

        Console.WriteLine();
    }

    Console.WriteLine();

    // Stop timing
    stopwatch.Stop();
    Console.WriteLine("Time taken : {0}", stopwatch.Elapsed);
}

static void PrintRowDataAndGameRowData(GridData gridData, string gameDate)
{
    foreach (var rowData in gridData.RowData)
    {
        var rowValues = rowData.Values.Select(cellData => cellData.FormattedValue);
        Console.WriteLine(String.Join(", ", rowValues));

        var gameRowData = GameRowData.LoadFromRowData(rowData, gameDate);
        Console.WriteLine(gameRowData.Player);
        Console.WriteLine(string.Join(",", gameRowData.GamesWon));
        Console.WriteLine(string.Join(",", gameRowData.GamesLost));
    }
}

static void PrintGames(string gameDate, GridData gridData)
{
    var gameGridData = GameGridData.LoadFromGridData(gridData, gameDate);

    foreach (var game in gameGridData.Games)
    {
        Console.WriteLine($"game.Id: {game.Id}");
        Console.WriteLine($"game.Winners: {string.Join(",", game.Winners)}");
        Console.WriteLine($"game.Losers: {string.Join(",", game.Losers)}");
        Console.WriteLine();
    }
}