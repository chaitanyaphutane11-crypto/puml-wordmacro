using System.Collections.Generic;
using System.IO;

namespace UmlToVbaMacroExtension
{
    public static class VbaMacroGenerator
    {
        public static void GenerateVba(List<UmlElement> elements, string outputPath, MacroOptionsPage options)
        {
            using (StreamWriter sw = new StreamWriter(outputPath))
            {
                sw.WriteLine("Option Explicit");
                sw.WriteLine();

                var enumEl = elements.Find(e => e.Type == "enum");
                if (enumEl != null)
                {
                    sw.WriteLine($"Public Enum {enumEl.Name}");
                    int i = 1;
                    foreach (var member in enumEl.Members)
                        sw.WriteLine($"    {member} = {i++}");
                    sw.WriteLine("End Enum");
                    sw.WriteLine();
                }

                sw.WriteLine($"Public Sub Execute(op As {enumEl.Name})");
                sw.WriteLine("    Select Case op");
                foreach (var member in enumEl.Members)
                {
                    if (member == "Cleanup" && !options.IncludeCleanup) continue;
                    if (member == "CustomLogic" && !options.IncludeCustomLogic) continue;
                    sw.WriteLine($"        Case {member}: Call {member}Func");
                }
                sw.WriteLine("    End Select");
                sw.WriteLine("End Sub");
                sw.WriteLine();

                foreach (var el in elements)
                {
                    if (el.Type == "macro")
                    {
                        foreach (var member in el.Members)
                        {
                            if (member == "Cleanup" && !options.IncludeCleanup) continue;
                            if (member == "CustomLogic" && !options.IncludeCustomLogic) continue;

                            sw.WriteLine($"Private Sub {member}Func()");
                            switch (member)
                            {
                                case "FindReplace":
                                    sw.WriteLine("    With Selection.Find");
                                    sw.WriteLine("        .Text = \": \"");
                                    sw.WriteLine("        .Replacement.Text = \":^t\"");
                                    sw.WriteLine("        .Forward = True");
                                    sw.WriteLine("        .Wrap = wdFindContinue");
                                    sw.WriteLine("    End With");
                                    sw.WriteLine("    Selection.Find.Execute Replace:=wdReplaceAll");
                                    break;
                                case "ConvertToTable":
                                    sw.WriteLine("    Selection.Range.ConvertToTable Separator:=wdSeparateByTabs");
                                    break;
                                case "FormatTable":
                                    sw.WriteLine("    Dim tbl As Table");
                                    sw.WriteLine("    If Selection.Tables.Count > 0 Then");
                                    sw.WriteLine("        Set tbl = Selection.Tables(1)");
                                    sw.WriteLine("        tbl.Columns(1).Range.Font.Bold = True");
                                    sw.WriteLine("        tbl.Rows(1).Shading.BackgroundPatternColor = wdColorGray20");
                                    sw.WriteLine("        tbl.Rows(1).Range.Font.Bold = True");
                                    sw.WriteLine("        tbl.Borders.Enable = True");
                                    sw.WriteLine("    End If");
                                    break;
                                case "Cleanup":
                                    sw.WriteLine("    With Selection.Find");
                                    sw.WriteLine("        .Text = \"^p^p\"");
                                    sw.WriteLine("        .Replacement.Text = \"^p\"");
                                    sw.WriteLine("    End With");
                                    sw.WriteLine("    Selection.Find.Execute Replace:=wdReplaceAll");
                                    break;
                                case "CustomLogic":
                                    sw.WriteLine("    MsgBox \"Running custom macro logic...\"");
                                    break;
                            }
                            sw.WriteLine("End Sub");
                            sw.WriteLine();
                        }
                    }
                }
            }
        }
    }
}
