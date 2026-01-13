

namespace GoContactSyncMod
{
    public static class StringExtensions
    {
        public static string Truncate(this string s, int maxLength, string suffix = "...")
        {
            if (maxLength > 0)
            {
                var length = maxLength - suffix.Length;
                if (length <= 0)
                {
                    return s;
                }
                if ((s != null) && (s.Length > maxLength))
                {
                    return s.Substring(0, length).TrimEnd(new char[0]) + suffix;
                }
            }
            return s;
        }

        public static string RemoveNewLines(this string s)
        {
            if (s == null)
            {
                return s;
            }
            else
            {
                return s.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
            }
        }
    }
}
