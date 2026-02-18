using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.IO;

namespace UmlToVbaMacroExtension
{
    public class MacroOptionsPage : DialogPage
    {
        [Category("Macro Generator")]
        [DisplayName("Output Folder")]
        [Description("Folder where the generated VBA macro file (.bas) will be saved.")]
        public string OutputFolder { get; set; } = @"C:\Temp";

        [Category("Macro Generator")]
        [DisplayName("Include Cleanup Function")]
        [Description("If true, the Cleanup function will be generated.")]
        public bool IncludeCleanup { get; set; } = true;

        [Category("Macro Generator")]
        [DisplayName("Include Custom Logic Function")]
        [Description("If true, the CustomLogic function will be generated.")]
        public bool IncludeCustomLogic { get; set; } = true;

        public override void SaveSettingsToStorage()
        {
            if (!Directory.Exists(OutputFolder))
            {
                throw new InvalidOperationException(
                    $"The folder '{OutputFolder}' does not exist. Please choose a valid folder."
                );
            }
            base.SaveSettingsToStorage();
        }
    }
}
