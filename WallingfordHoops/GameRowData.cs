using Google.Apis.Sheets.v4.Data;

namespace WallingfordHoops
{
    internal class GameRowData
    {
        public GameRowData(string player)
        {
            Player = player;
            GamesWon = new List<string>();
            GamesLost = new List<string>();
        }
        public string Player {  get; set; }
        public IList<string> GamesWon {  get; set; }
        public IList<string> GamesLost { get; set; }

        public static GameRowData LoadFromRowData(RowData rowData, string gameDate)
        {
            var rowValues = rowData.Values.Select(cellData => cellData.FormattedValue).ToList();
            var player = rowValues[0];
            var gameRowData = new GameRowData(player);

            for (int i = 3; i < rowValues.Count; i++)
            {
                var gameId = $"{gameDate}.{i - 2}";
                if (rowValues[i] == "W")
                {
                    gameRowData.GamesWon.Add(gameId);
                }
                if (rowValues[i] == "L")
                {
                    gameRowData.GamesLost.Add(gameId);
                }
            }

            return gameRowData;
        }

        public static bool IsLoadable(RowData rowData)
        {
            var winsColumn = rowData.Values[1].FormattedValue;
            var lossesColumn = rowData.Values[2].FormattedValue;

            if (winsColumn == null || winsColumn == "W") { return false; }
            if (lossesColumn == null || lossesColumn == "L") { return false; }
            if (winsColumn == "0" && lossesColumn == "0") { return false; }
            return true;
        }
    }
}
