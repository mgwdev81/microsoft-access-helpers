using Microsoft.Office.Interop.Access;

namespace ApplicationOption
{
    public static class ApplicationOption
    {
        public static void SetOption(Application application, string optionName, string optionValue)
        {
            application.SetOption(optionName, optionValue);
        }

        public static dynamic GetOption(Application application, string optionName)
        {
            return application.GetOption(optionName);
        }
    }
}