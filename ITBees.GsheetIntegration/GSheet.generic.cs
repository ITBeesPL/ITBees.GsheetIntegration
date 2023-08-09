using GSheetConnector;

namespace ITBees.GsheetIntegration
{
    public class GSheet<T> : GSheet
    {
        public new readonly IList<T> Data;

        public GSheet(IList<T> data, Dictionary<int, string> columnsIndexes)
        {
            ColumnsIndexes = columnsIndexes;
            Data = data;
        }
    }
}