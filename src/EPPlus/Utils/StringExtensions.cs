namespace OfficeOpenXml.Utils;

internal static class StringExtensions
{
    internal static string NullIfWhiteSpace(this string s) => s == "" ? null : s;
}