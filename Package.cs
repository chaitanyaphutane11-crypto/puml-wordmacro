using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace UmlToVbaMacroExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("UML to VBA Macro", "Generates Word VBA macros from PlantUML", "1.0")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideOptionPage(typeof(MacroOptionsPage), "Macro Generator", "General", 0, 0, true)]
    [Guid(PackageGuidString)]
    public sealed class UmlToVbaMacroExtensionPackage : AsyncPackage
    {
        public const string PackageGuidString = "d6a2f5b1-1234-4a56-9876-abcdef123456";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await GenerateMacroCommand.InitializeAsync(this);
        }
    }
}
