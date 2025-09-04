using System;
using System.IO;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    public class GH_Calcpad_Export_word : GH_Component
    {
        public GH_Calcpad_Export_word()
          : base("Export Word", "ExportWord",
                 "Exports Calcpad to .docx using Calcpad CLI (OpenXml). Fallback: Microsoft Word (OMath).",
                 "Calcpad", "6. Saving & Export")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Updated Sheet", "US", "Calculated CalcpadSheet (from Play CPD)", GH_ParamAccess.item);
            p.AddTextParameter("File", "N", "Base name (without extension)", GH_ParamAccess.item);
            p.AddTextParameter("Output Folder", "F", "Destination folder", GH_ParamAccess.item);
            p.AddBooleanParameter("Execute", "X", "True = export", GH_ParamAccess.item, false);

            p[1].Optional = true;
            p[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("Word Path", "W", "Path of the generated .docx", GH_ParamAccess.item);
            p.AddBooleanParameter("Success", "S", "True if .docx was generated", GH_ParamAccess.item);
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
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Pon Execute=True para exportar Word");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Conecta 'Output Folder'. No se exporta hasta tener ruta.");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Conecta 'File'. No se exporta hasta tener nombre.");
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
            if (string.IsNullOrWhiteSpace(sheet.OriginalCode))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No code to export. Ejecuta 'Play CPD' antes.");
                return;
            }

            // 1) Intentar CLI (resultado idéntico a Calcpad)
            var res = CalcpadExporter.ExportDocxCli(
                sheet,
                outputFolder,
                fileName.Trim(),
                cliPath: null, // usa autodetección (o pasa una ruta aquí)
                msg => AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, msg));

            // 2) Fallback a Interop si la CLI no está disponible
            if (!res.success)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "CLI no disponible o falló. Intentando Word Interop…");
                res = CalcpadExporter.ExportDocxInterop(
                    sheet,
                    outputFolder,
                    fileName.Trim(),
                    msg => AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, msg));
            }

            DA.SetData(0, res.success ? res.finalPath : null);
            DA.SetData(1, res.success);

            if (res.success)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"✅ DOCX export → {Path.GetFileName(res.finalPath)} | {res.size / 1024.0:F1} KB | {res.method}");
            else
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "❌ DOCX export failed. Configura la CLI de Calcpad o revisa el Interop de Word.");
        }

        public override Guid ComponentGuid => new Guid("C3F2D4E5-F6A7-4B8C-9D0E-1F2A3B4C5D6F");
        protected override Bitmap Icon => Resources.Icon_Word;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
