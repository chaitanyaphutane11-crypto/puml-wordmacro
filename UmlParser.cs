using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace UmlToVbaMacroExtension
{
    public class UmlElement
    {
        public string Name { get; set; }
        public string Type { get; set; } // "class", "enum", "macro"
        public List<string> Members { get; set; } = new List<string>();
    }

    public static class UmlParser
    {
        public static List<UmlElement> Parse(string[] lines)
        {
            var elements = new List<UmlElement>();
            UmlElement current = null;

            foreach (var line in lines)
            {
                var classMatch = Regex.Match(line, @"class\s+(\w+)\s*(<<(\w+)>>)?");
                if (classMatch.Success)
                {
                    current = new UmlElement
                    {
                        Name = classMatch.Groups[1].Value,
                        Type = classMatch.Groups[3].Success ? classMatch.Groups[3].Value.ToLower() : "class"
                    };
                    elements.Add(current);
                    continue;
                }

                var enumMatch = Regex.Match(line, @"enum\s+(\w+)\s*(<<(\w+)>>)?");
                if (enumMatch.Success)
                {
                    current = new UmlElement
                    {
                        Name = enumMatch.Groups[1].Value,
                        Type = "enum"
                    };
                    elements.Add(current);
                    continue;
                }

                var memberMatch = Regex.Match(line, @"([+\-#])?(\w+)\(?\)?");
                if (memberMatch.Success && current != null)
                {
                    current.Members.Add(memberMatch.Groups[2].Value);
                }
            }

            return elements;
        }
    }
}
