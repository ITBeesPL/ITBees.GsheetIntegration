using System.Dynamic;
using System.Reflection;
using Google.Apis.Sheets.v4.Data;
using ITBees.GsheetIntegration.Interfaces;
using ITBees.Models.Languages;
using ITBees.Translations;

namespace ITBees.GsheetIntegration.Tools
{
    public class GSheetTypeConverter
    {
        public static GSheet Convert(ValueRange valueRange, bool firstRowHeader)
        {
            var results = new List<IDictionary<string, object>>();

            var columns = CreateDictionaryOfColumns(valueRange, firstRowHeader);

            for (int i = 1; i < valueRange.Values.Count; i++)
            {
                var row = valueRange.Values[i];

                var instance = new ExpandoObject() as IDictionary<string, Object>;

                if (CheckIsNotEmptyRow(row))
                {
                    for (int j = 0; j < row.Count; j++)
                    {
                        var propertyName = columns.Where(x => x.Key == j).First().Value;
                        //instance.Keys.Add(propertyName);
                        try
                        {
                            var count = instance.Keys.Where(x => x.StartsWith(propertyName)).Count();
                            if (count == 0)
                            {
                                instance.Add(propertyName, row[j]);
                            }
                            else
                            {
                                instance.Add(propertyName + (count + 1), row[j]);
                            }

                        }
                        catch (Exception e)
                        {
                            throw new Exception(
                                $"Check collumn name :{propertyName} with value {row[j]} error message : {e.Message}");
                        }

                    }

                    results.Add(instance);
                }
            }

            return new GSheet(results, columns);
        }

        public static GSheet<T> Convert<T>(ValueRange valueRange, bool firstRowHeader, Language lang) where T : class, IGuidItem
        {
            var results = new List<T>();

            var columns = CreateDictionaryOfColumns(valueRange, firstRowHeader);

            for (int i = 1; i < valueRange.Values.Count; i++)
            {
                var row = valueRange.Values[i];

                T instance = Activator.CreateInstance<T>();
                instance.WorksheetName = typeof(T).GetType().Name;
                if (CheckIsNotEmptyRow(row))
                {

                    for (int j = 0; j < row.Count; j++)
                    {
                        var propertyName = columns.Where(x => x.Key == j).First().Value;

                        var currentValue = row[j];
                        if (propertyName.StartsWith("Guid"))
                        {
                            PropertyInfo pi = instance.GetType().GetProperty(propertyName.Replace("-", ""));
                            if (currentValue.ToString().Length < 32)
                            {
                                throw new NotProperGuidValueInsideGsheetException(currentValue.ToString(), lang);
                            }
                            if (string.IsNullOrEmpty(currentValue.ToString()) == false)
                            {
                                pi.SetValue(instance, new Guid(currentValue.ToString()));
                            }
                        }
                        else
                        {
                            PropertyInfo pi = instance.GetType().GetProperty(propertyName.Replace("-", "") + "_" + j);
                            if (pi.PropertyType == typeof(string))
                            {
                                pi.SetValue(instance, currentValue.ToString());
                            }
                            if (pi.PropertyType == typeof(bool))
                            {
                                pi.SetValue(instance, bool.Parse(currentValue.ToString()));
                            }
                            if (pi.PropertyType == typeof(DateTime?))
                            {
                                if (currentValue.ToString() == "")
                                {
                                    currentValue = null;
                                }
                                else
                                {
                                    pi.SetValue(instance, DateTime.Parse(currentValue.ToString()));
                                }
                            }
                        }
                    }

                    results.Add(instance);
                }
            }

            return new GSheet<T>(results, columns);

        }

        private static Dictionary<int, string> CreateDictionaryOfColumns(ValueRange valueRange, bool firstRowHeader)
        {
            if (firstRowHeader == false)
            {
                throw new Exception("Unable to automaticaly convert data if first row is provided");
            }

            var firstRow = valueRange.Values.First();
            var dict = new Dictionary<int, string>();
            for (int i = 0; i < firstRow.Count; i++)
            {
                dict.Add(i, firstRow[i].ToString().Trim());
            }

            return dict;
        }

        private static bool CheckIsNotEmptyRow(IList<object> row)
        {
            foreach (object value in row)
            {
                if (value.ToString().Contains("FALSE") || value.ToString().Contains("TRUE"))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(value.ToString()) == false)
                {
                    return true;
                }
            }

            return false;
        }
    }

    public class NotProperGuidValueInsideGsheetException : Exception
    {
        public NotProperGuidValueInsideGsheetException(string message, Language lang) : base($"{Translate.Get(() => Translations.GoogleSheetContectorTransaltions.NotProperGuidValueInsideGsheetException, lang)}{message}")
        {

        }
    }
}