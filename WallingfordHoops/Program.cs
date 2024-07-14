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
        if (sheet.Properties.Title.StartsWith("202"))
        {
            rangesForSheetsWithGameData.Add($"{sheet.Properties.Title}!A1:O17");
        }
    }

    Console.WriteLine();

    getSpreadsheetRequest.IncludeGridData = true;
    getSpreadsheetRequest.Ranges = rangesForSheetsWithGameData;
    var spreadsheetWithGameData = getSpreadsheetRequest
        .Execute();

    // PrintGamesFromSpreadsheet(spreadsheetWithGameData);
    // PrintRowDataAndGameRowDataFromSpreadsheet(spreadsheetWithGameData);

    PrintHenryStatsFromSpreadsheet(spreadsheetWithGameData);

    Console.WriteLine();

    PrintMoreStatsFromSpreadsheet(spreadsheetWithGameData);

    // Stop timing
    stopwatch.Stop();
    Console.WriteLine("Time taken : {0}", stopwatch.Elapsed);
}

static IList<Game> GetGamesFromSpreadsheet(Spreadsheet spreadsheet)
{
    var games = new List<Game>();
    foreach (var sheet in spreadsheet.Sheets)
    {
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];

        var gameGridData = GameGridData.LoadFromGridData(gridData, gameDate);
        games.AddRange(gameGridData.Games);
    }

    return games;
}

static IList<string> GetPlayersFromSpreadsheet(Spreadsheet spreadsheet)
{
    var players = new SortedSet<string>();
    foreach (var sheet in spreadsheet.Sheets)
    {
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];

        var gameGridData = GameGridData.LoadFromGridData(gridData, gameDate);
        foreach (var player in gameGridData.Players)
        {
            players.Add(player);
        }
    }

    return players.ToList();
}

static void PrintRowDataAndGameRowDataFromSpreadsheet(Spreadsheet spreadsheet)
{
    Console.WriteLine("Printing all row data and game row data:");
    foreach (var sheet in spreadsheet.Sheets)
    {
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];
        Console.WriteLine(gameDate);

        PrintRowDataAndGameRowData(gridData, gameDate);

        Console.WriteLine();
    }
}

static void PrintGamesFromSpreadsheet(Spreadsheet spreadsheet)
{
    Console.WriteLine("Printing all games:");
    foreach (var sheet in spreadsheet.Sheets)
    {
        var gameDate = sheet.Properties.Title;
        var gridData = sheet.Data[0];
        Console.WriteLine(gameDate);

        PrintGames(gameDate, gridData);

        Console.WriteLine();
    }
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

static void PrintMoreStatsFromSpreadsheet(Spreadsheet spreadsheet)
{
    var games = GetGamesFromSpreadsheet(spreadsheet);
    var players = GetPlayersFromSpreadsheet(spreadsheet);
    var weeks = games.Select(game => game.Id.Split('.')[0]).Distinct();

    Console.WriteLine($"Data current through: {weeks.Order().Last()}");
    Console.WriteLine($"Number of players: {players.Count}");
    Console.WriteLine($"Number of weeks: {weeks.Count()}");
    Console.WriteLine($"Number of games: {games.Count}");
    Console.WriteLine($"Player,WeeksPlayed,GamesPlayed");

    foreach (var player in players)
    {
        var playerGamesPlayed = from game in games
                                where game.Winners.Contains(player) || game.Losers.Contains(player)
                                select game;
        var playerWeeksPlayed = playerGamesPlayed.Select(game => game.Id.Split('.')[0]).Distinct();
        Console.WriteLine($"{player},{playerWeeksPlayed.Count()},{playerGamesPlayed.Count()}");
    }
}

static void PrintHenryStatsFromSpreadsheet(Spreadsheet spreadsheet)
{
    var games = GetGamesFromSpreadsheet(spreadsheet);

    var henryGamesWon = from game in games
                        where game.Winners.Contains("Henry")
                        select game;

    var henryGamesLost = from game in games
                         where game.Losers.Contains("Henry")
                         select game;

    var henryWinRate = 1.0 * henryGamesWon.Count() / (henryGamesWon.Count() + henryGamesLost.Count());

    var henryGamesPlayed = from game in games
                           where game.Winners.Contains("Henry") || game.Losers.Contains("Henry")
                           select game;

    var henryWeeksPlayed = henryGamesPlayed.Select(game => game.Id.Split('.')[0]).Distinct();

    Console.WriteLine($"Henry's wins: {henryGamesWon.Count()}");
    Console.WriteLine($"Henry's losses: {henryGamesLost.Count()}");
    Console.WriteLine($"Henry's win rate: {Math.Round(henryWinRate, 3)}");
    Console.WriteLine($"Henry's henryWeeksPlayed: {henryWeeksPlayed}");
    foreach (var week in henryWeeksPlayed)
    {
        Console.WriteLine(week);
    }
    Console.WriteLine($"Henry's henryWeeksPlayed: {henryWeeksPlayed}");


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
}