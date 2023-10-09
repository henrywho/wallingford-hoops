namespace WallingfordHoops
{
    internal class Game
    {
        public Game(string id)
        {
            Id = id;
            Winners = new List<string>();
            Losers = new List<string>();
        }

        public string Id { get; }
        public IList<string> Winners { get; set; }
        public IList<string> Losers { get; set; }
    }
}
