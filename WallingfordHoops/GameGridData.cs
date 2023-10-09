using Google.Apis.Sheets.v4.Data;

namespace WallingfordHoops
{
    internal class GameGridData
    {
        public GameGridData()
        {
            Games = new List<Game>();
        }

        public IList<Game> Games { get; set; }

        public void InsertWinnerIntoGame(string player, string gameId)
        {
            var game = GetOrCreateGame(gameId);
            game.Winners.Add(player);
        }
        public void InsertLoserIntoGame(string player, string gameId)
        {
            var game = GetOrCreateGame(gameId);
            game.Losers.Add(player);
        }

        public Game GetOrCreateGame(string gameId)
        {
            var game = Games.SingleOrDefault(game => game.Id == gameId);
            if (game == null)
            {
                game = new Game(gameId);
                Games.Add(game);
            }
            return game;
        }

        public static GameGridData LoadFromGridData(GridData gridData, string gameDate)
        {
            var gameGridData = new GameGridData();

            foreach (var rowData in gridData.RowData)
            {
                if (!GameRowData.IsLoadable(rowData))
                {
                    continue;
                }

                var gameRowData = GameRowData.LoadFromRowData(rowData, gameDate);
                foreach (var gameId in gameRowData.GamesWon)
                {
                    gameGridData.InsertWinnerIntoGame(gameRowData.Player, gameId);
                }
                foreach (var gameId in gameRowData.GamesLost)
                {
                    gameGridData.InsertLoserIntoGame(gameRowData.Player, gameId);
                }
            }

            return gameGridData;
        }
    }
}
