using System;
using System.IO;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Exports Calcpad's own rendered HTML report (same worksheet template Calcpad's
    /// desktop app/CLI use).
    /// </summary>
    public class GH_Calcpad_Export_html : GH_Component
    {
        public GH_Calcpad_Export_html()
          : base(
                "Export HTML",
                "ExportHTML",
                "Exports Calcpad's rendered HTML report",
                "Calcpad",
                "6. Saving & Export")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "CalcpadSheet to export (from Play CPD or Search Variables)", GH_ParamAccess.item);
            p.AddTextParameter("File", "N", "Base name (without extension)", GH_ParamAccess.item);
            p.AddTextParameter("Output Folder", "F", "Destination folder", GH_ParamAccess.item);
            p.AddBooleanParameter("Execute", "X", "True = export", GH_ParamAccess.item, false);

            p[1].Optional = true;
            p[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("HTML Path", "H", "Path of the generated .html", GH_ParamAccess.item);
            p.AddBooleanParameter("Success", "OK", "True if .html was generated", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object data = null;
            string fileName = null;
            string outputFolder = null;
            bool execute = false;

            if (!DA.GetData(0, ref data)) return;
            DA.GetData(1, ref fileName);
            DA.GetData(2, ref outputFolder);
            DA.GetData(3, ref execute);

            if (!execute)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Set Execute=True to export HTML");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            var sheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid CalcpadSheet.");
                return;
            }

            var (success, finalPath, error) = CalcpadExporter.Export(sheet, outputFolder, fileName, "html");

            DA.SetData(0, success ? finalPath : null);
            DA.SetData(1, success);

            if (success)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"HTML export -> {Path.GetFileName(finalPath)}");
            else
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"HTML export failed: {error}");
        }

        public override Guid ComponentGuid => new Guid("0D1D3B68-9D8B-4F6A-A23E-6F7EB07E2E7D");
        protected override Bitmap Icon => Resources.Icon_Html;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
