using System;
using System.Globalization;
using System.Text;

namespace TaxPersonnelManagement.Helpers
{
    public static class SearchHelper
    {
        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    // Special treatment for 'Đ' and 'đ'
                    if (c == 'Đ')
                        stringBuilder.Append('D');
                    else if (c == 'đ')
                        stringBuilder.Append('d');
                    else
                        stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        public static bool IsMatch(string? source, string keyword)
        {
            if (string.IsNullOrEmpty(keyword))
                return true;

            if (string.IsNullOrEmpty(source))
                return false;

            string cleanSource = RemoveDiacritics(source);
            string cleanKeyword = RemoveDiacritics(keyword);

            return cleanSource.Contains(cleanKeyword, StringComparison.OrdinalIgnoreCase);
        }
    }
}
