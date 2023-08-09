namespace ITBees.GsheetIntegration.Models;

public class GoogleSheetRow
{
    public GoogleSheetRow() => Cells = new List<GoogleSheetCell>();

    public List<GoogleSheetCell> Cells { get; set; }
}