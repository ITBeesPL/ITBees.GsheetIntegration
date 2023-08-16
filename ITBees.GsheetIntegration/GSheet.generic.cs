using ITBees.GsheetIntegration.Interfaces;

namespace ITBees.GsheetIntegration
{
    public class GSheet<T> : GSheet where T : class, IGuidItem
    {
        public new readonly IList<T> Data;

        public GSheet(IList<T> data, Dictionary<int, string> columnsIndexes)
        {
            ColumnsIndexes = columnsIndexes;
            Data = data;
        }
    }
}