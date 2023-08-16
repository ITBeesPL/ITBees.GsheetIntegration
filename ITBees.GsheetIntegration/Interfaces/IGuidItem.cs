namespace ITBees.GsheetIntegration.Interfaces;

public interface IGuidItem
{
    public Guid Guid { get; set; }
    public string WorksheetName { get; set; }
}