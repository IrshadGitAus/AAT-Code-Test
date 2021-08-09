using System.Text.RegularExpressions;

namespace AAT.File.Processor3.Helpers
{
    public static class ApplicationHelper
    {
        public static string GetApplicationRoot()
        {
            var exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Regex appPathMatcher = new Regex(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = appPathMatcher.Match(exePath).Value;
            return appRoot;
        }
    }
}
