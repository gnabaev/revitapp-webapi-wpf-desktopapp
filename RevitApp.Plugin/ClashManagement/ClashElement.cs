namespace RevitApp.Plugin.ClashManagement
{
    public class ClashElement
    {
        public string Clash { get; set; }

        public int Id { get; set; }

        public string Model { get; set; }

        public ClashElement(string clash, int id, string model)
        {
            Clash = clash;
            Id = id;
            Model = model;
        }
    }
}
