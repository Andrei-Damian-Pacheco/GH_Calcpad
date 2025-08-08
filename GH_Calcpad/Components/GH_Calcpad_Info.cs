using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Drawing;
using Grasshopper.Kernel;
using Calcpad.Core;

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

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Plugin info
            var pluginVersion = Assembly
                .GetExecutingAssembly()
                .GetName()
                .Version
                .ToString();
            var pluginInfo = $"GH_Calcpad v{pluginVersion}";

            // 2) Calcpad.Core.dll info: version + license
            string coreVersion = "Unknown";
            // Fallback: if we don't find the file, show these two lines
            string licenseText =
                "MIT License" + Environment.NewLine +
                "Copyright (c) 2023 Proektsoft EOOD";

            try
            {
                // Get the Calcpad.Core assembly
                var coreAsm = typeof(Settings).Assembly;
                coreVersion = coreAsm.GetName().Version.ToString();

                // Try to read doc\LICENSE.TXT next to the DLL
                var baseDir = Path.GetDirectoryName(coreAsm.Location);
                var licensePath = Path.Combine(baseDir, "doc", "LICENSE.TXT");

                if (File.Exists(licensePath))
                {
                    // Read only the first two real lines
                    var firstTwo = File
                        .ReadLines(licensePath)
                        .Take(2);
                    licenseText = string.Join(Environment.NewLine, firstTwo);
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    $"Error reading Calcpad.Core info: {ex.Message}"
                );
            }

            // Combine everything into a single string
            var calcpadInfo =
                $"Calcpad.Core v{coreVersion}{Environment.NewLine}{licenseText}";

            // Set outputs
            DA.SetData(0, pluginInfo);
            DA.SetData(1, calcpadInfo);
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override Bitmap Icon
            => Properties.Resources.Icon_Calcpad;

        public override Guid ComponentGuid
            => new Guid("64C4211A-79A2-48A0-9A2E-7CCF7ED6034E");
    }
}