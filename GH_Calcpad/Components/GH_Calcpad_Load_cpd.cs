using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Grasshopper.Kernel;
using GH_Calcpad.Classes;    // CalcpadSheet, CalcpadSyntax
using GH_Calcpad.Properties;  // Resources (icon)

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Lee un .cpd/.txt (código Calcpad), extrae variables/valores/unidades
    /// usando un parser basado en la sintaxis oficial (XML/Notepad++) con fallback.
    /// </summary>
    public class GH_Calcpad_Load_cpd : GH_Component
    {
        private FileSystemWatcher _watcher;

        public GH_Calcpad_Load_cpd()
          : base("Load CPD", "LoadCPD",
                 "Reads a .cpd file and extracts variables, values and units",
                 "Calcpad", "2. File Loading")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter("FilePath", "P", "Complete path to .cpd (or .txt) file", GH_ParamAccess.item);
            p.AddBooleanParameter("CaptureExplicit", "C", "If True, only captures explicit variables: 'var=?{val}unit' or 'valunit';'var'", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("Variables", "N", "Names of all variables found", GH_ParamAccess.list);
            p.AddNumberParameter("Values", "V", "Numeric values associated 1:1 with Variables", GH_ParamAccess.list);
            p.AddTextParameter("Units", "U", "Units corresponding 1:1 with Variables", GH_ParamAccess.list);
            p.AddGenericParameter("SheetObj", "S", "CalcpadSheet instance for later consumption", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string path = null;
            bool captureExplicit = false;
            if (!DA.GetData(0, ref path)) return;
            DA.GetData(1, ref captureExplicit);

            SetupFileWatcher(path);

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"File does not exist:\n{path}");
                return;
            }

            string content;
            try
            {
                content = File.ReadAllText(path, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Could not read file: {ex.Message}");
                return;
            }

            try
            {
                // Nuevo parser basado en sintaxis
                CalcpadSyntax.Instance.ParseVariables(content, captureExplicit, out var names, out var values, out var units);

                var sheet = new CalcpadSheet(names, values, units);
                try { sheet.SetFullCode(content); } catch { }

                if (names.Count != values.Count || values.Count != units.Count)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Mismatch in Variables/Values/Units");

                DA.SetDataList(0, names);
                DA.SetDataList(1, values);
                DA.SetDataList(2, units);
                DA.SetData(3, sheet);

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"Loaded {names.Count} variable(s) | ExplicitMode={captureExplicit}");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Parse error: {ex.Message}");
            }
        }

        private void SetupFileWatcher(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
                string file = Path.GetFileName(filePath);
                if (_watcher != null && _watcher.Path == dir && _watcher.Filter == file)
                    return;

                if (_watcher != null)
                {
                    _watcher.Dispose();
                    _watcher = null;
                }

                _watcher = new FileSystemWatcher(dir, file)
                {
                    NotifyFilter = NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };
                _watcher.Changed += (s, e) => ExpireSolution(true);
            }
            catch { }
        }

        public override void RemovedFromDocument(Grasshopper.Kernel.GH_Document document)
        {
            try
            {
                _watcher?.Dispose();
                _watcher = null;
            }
            catch { /* ignore */ }

            base.RemovedFromDocument(document);
        }

        public override Guid ComponentGuid => new Guid("7A4CE2C1-4F7D-4C7E-A5E1-5B0C2F7E8F13");
        protected override System.Drawing.Bitmap Icon => Resources.Icon_Calc;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}