using System;
using System.Linq;

namespace ProcessTimesheets
{
    static class FilenameCleaner
    {
        private static readonly char[] illegalChars = System.IO.Path.GetInvalidFileNameChars();

        public static string cleanFilename(string maybeDirty, string replacement = "_")
        {
            return string.Join(replacement, maybeDirty.Split(illegalChars));
        }
    }
}
