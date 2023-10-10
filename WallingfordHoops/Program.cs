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
        .Get(configuration["WallingfordHoopsSpreadsheetId"]);
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

    var games = GetAllGamesFromSpreadsheet(sheetsWithGameData);

    var henryGamesWon = from game in games
                        where game.Winners.Contains("Henry")
                        select game;

    var henryGamesLost = from game in games
                         where game.Losers.Contains("Henry")
                         select game;

    var henryWinRate = 1.0 * henryGamesWon.Count() / (henryGamesWon.Count() + henryGamesLost.Count());

    Console.WriteLine($"Henry's wins: {henryGamesWon.Count()}");
    Console.WriteLine($"Henry's losses: {henryGamesLost.Count()}");
    Console.WriteLine($"Henry's win rate: {Math.Round(henryWinRate, 3)}");

    Console.WriteLine();

    var henryVsGelGamesWon = from game in games
                             where game.Winners.Contains("Henry")
                             where game.Losers.Contains("Gel")
                             select game;

    var henryVsGelGamesLost = from game in games
                              where game.Losers.Contains("Henry")
                              where game.Winners.Contains("Gel")
                              select game;

    var henryVsGelWinRate = 1.0 * henryVsGelGamesWon.Count() / (henryVsGelGamesWon.Count() + henryVsGelGamesLost.Count());

    Console.WriteLine($"Henry vs Gel wins: {henryVsGelGamesWon.Count()}");
    Console.WriteLine($"Henry vs Gel losses: {henryVsGelGamesLost.Count()}");
    Console.WriteLine($"Henry vs Gel win rate: {Math.Round(henryVsGelWinRate, 3)}");

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
        Console.WriteLine(string.Join(", ", rowValues));

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

static IList<Game> GetAllGamesFromSpreadsheet(IList<Sheet> spreadsheet)
{
    var games = new List<Game>();
    foreach (var sheet in spreadsheet)
    {
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];

        var gameGridData = GameGridData.LoadFromGridData(gridData, gameDate);
        games.AddRange(gameGridData.Games);
    }

    return games;
}