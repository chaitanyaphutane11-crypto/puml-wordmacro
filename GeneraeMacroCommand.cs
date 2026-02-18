using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Threading.Tasks;

namespace UmlToVbaMacroExtension
{
    internal sealed class GenerateMacroCommand
    {
        public static async Task ExecuteAsync(string pumlPath)
        {
            var options = (MacroOptionsPage)await UmlToVbaMacroExtensionPackage
                .GetGlobalServiceAsync(typeof(MacroOptionsPage));

            string outputPath = Path.Combine(options.OutputFolder, "WordMacro.bas");

            var elements = UmlParser.Parse(File.ReadAllLines(pumlPath));
            VbaMacroGenerator.GenerateVba(elements, outputPath, options);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Hook up menu command here if desired
        }
    }
}
