using System.Collections;

namespace ITBees.GsheetIntegration
{
    public class GSheet
    {
        public readonly IList Data;
        public Dictionary<int, string> ColumnsIndexes { get; protected set; }
        public GSheet(IList data, Dictionary<int, string> columnsIndexes)
        {
            ColumnsIndexes = columnsIndexes;
            Data = data;
        }

        public GSheet()
        {
            
        }
    }
}