using ITBees.GsheetIntegration.Interfaces;

namespace ITBees.GsheetIntegration.Models;

public abstract class WorksheetRecord : IGuidItem
{
    public string WorksheetName { get; set; }

    public WorksheetRecord(string worksheetName)
    {
        WorksheetName = worksheetName;
    }

    protected WorksheetRecord()
    {
        
    }

    public Guid Guid { get; set; }
}