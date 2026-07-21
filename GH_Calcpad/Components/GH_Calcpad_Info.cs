using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Drawing;
using Grasshopper.Kernel;

namespace GH_Calcpad.Components
{
    public class GH_Calcpad_Info : GH_Component
    {
        public GH_Calcpad_Info()
          : base(
              "Calcpad Info",    // Name
              "CP_Info",         // Nickname
              "Shows plugin version and Calcpad.Core version + license",
              "Calcpad",         // Tab
              "1. Information & Diagnostics"  // Info
          )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p) { }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("PluginInfo", "Plugin", "GH_Calcpad plugin version", GH_ParamAccess.item);
            p.AddTextParameter("CalcpadInfo", "Calcpad", "Calcpad.Core version and license", GH_ParamAccess.item);
        }

        // Computed once per Rhino session (no inputs, output never changes while the plugin is loaded)
        // and shared across all instances of this component to avoid repeated reflection/disk I/O on every solve.
        private static readonly Lazy<string> _pluginInfo = new Lazy<string>(BuildPluginInfo);
        private static readonly Lazy<string> _calcpadInfo = new Lazy<string>(BuildCalcpadInfo);

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            DA.SetData(0, _pluginInfo.Value);
            DA.SetData(1, _calcpadInfo.Value);
        }

        private static string BuildPluginInfo()
        {
            var asm = Assembly.GetExecutingAssembly();
            var pluginVersion = asm.GetName().Version.ToString();
            var author = asm.GetCustomAttribute<AssemblyCompanyAttribute>()?.Company;
            return string.IsNullOrWhiteSpace(author)
                ? $"GH_Calcpad v{pluginVersion}"
                : $"GH_Calcpad v{pluginVersion}{Environment.NewLine}{author}";
        }

        private static string BuildCalcpadInfo()
        {
            // Calcpad.Core is vendored source (Worker\Calcpad.Core), compiled into the
            // worker process rather than referenced in-process - so its version/authorship
            // is read straight off the worker's own Calcpad.Core.dll file metadata instead
            // of an assembly reference.
            try
            {
                var pluginDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
                var corePath = Path.Combine(pluginDir, "Worker", "Calcpad.Core.dll");
                if (File.Exists(corePath))
                {
                    var info = FileVersionInfo.GetVersionInfo(corePath);
                    var authorLine = string.IsNullOrWhiteSpace(info.LegalCopyright)
                        ? info.CompanyName
                        : $"{info.CompanyName} - {info.LegalCopyright}";
                    return $"Calcpad.Core v{info.FileVersion}{Environment.NewLine}{authorLine}";
                }
            }
            catch
            {
                // Diagnostic-only info; fall through to the "unavailable" message.
            }
            return "Calcpad.Core: version unavailable (Worker\\Calcpad.Core.dll not found)";
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override Bitmap Icon
            => Properties.Resources.Icon_Calcpad;

        public override Guid ComponentGuid
            => new Guid("64C4211A-79A2-48A0-9A2E-7CCF7ED6034E");
    }
}