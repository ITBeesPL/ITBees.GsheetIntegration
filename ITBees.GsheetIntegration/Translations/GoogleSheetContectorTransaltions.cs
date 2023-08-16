using ITBees.Translations.Interfaces;

namespace ITBees.GsheetIntegration.Translations;

public class GoogleSheetContectorTransaltions : ITranslate
{
    public static readonly string DefaultNewGoogleSheetName = "New worksheet";
    public static readonly string NotProperGuidValueInsideGsheetException = "Nieprawidłowa wartość Guid w arkuszu google - przykładowa postać to '08db9e54-edef-4a6a-8693-00004cb048c6";
}