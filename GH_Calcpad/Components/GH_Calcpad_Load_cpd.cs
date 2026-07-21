using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Grasshopper.Kernel;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Lee un .cpd/.txt, extrae variables (inputs), valores y unidades.
    /// Reordena las variables en el mismo orden del documento.
    /// </summary>
    public class GH_Calcpad_Load_cpd : GH_Component
    {
        private FileSystemWatcher _watcher;
        private volatile bool _frozen;

        // Compiladas una sola vez por tipo, en vez de reconstruirse en cada SolveInstance.
        private static readonly Regex RxSupplementalScan = new Regex(
            @"^(?<indent>\s*)(?<name>[A-Za-z_][A-Za-z0-9_′″‴⁗’\.,]*)\s*=\s*(?<num>[+-]?\d+(?:\.\d+)?)(?:\s*(?<unit>[A-Za-z°µμΩ℧·/\-\^²³]+(?:\/[A-Za-z°µμΩ℧·/\-\^²³]+)*))?\s*(?:$|[#'’‘])",
            RegexOptions.Multiline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex RxHasRightSideOperator = new Regex(
            @"=\s*[+-]?\d+(?:\.\d+)?\s*[\+\-\*/\^]", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex RxAssignment = new Regex(
            @"(?<![A-Za-z0-9_])(?<name>[A-Za-z_][A-Za-z0-9_′″‴⁗’\.,]*)\s*=",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public GH_Calcpad_Load_cpd()
          : base("Load CPD", "LoadCPD",
                 "Reads a .cpd file and extracts variables, values and units",
                 "Calcpad", "2. File Loading")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter("FilePath", "P", "Complete path to .cpd (or .txt) file", GH_ParamAccess.item);
            p.AddBooleanParameter("Freeze", "F", "Acts as a switch: if True, blocks all output (no data passes downstream) and stops reading/watching the file. Default = false (always reads on solve).", GH_ParamAccess.item, false);
            p[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("Variables", "N", "Names of variables found (ordered as in file)", GH_ParamAccess.list);
            p.AddNumberParameter("Values", "V", "Numeric values 1:1 with Variables", GH_ParamAccess.list);
            p.AddTextParameter("Units", "U", "Units 1:1 with Variables", GH_ParamAccess.list);
            p.AddGenericParameter("SheetObj", "S", "CalcpadSheet instance", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string path = null;
            bool freeze = false;
            if (!DA.GetData(0, ref path)) return;
            DA.GetData(1, ref freeze);
            _frozen = freeze;

            SetupFileWatcher(path);

            if (freeze)
            {
                // Acts as a switch, not a cache: no data passes through at all while frozen,
                // rather than replaying the last successfully loaded values.
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Freeze=True: reading paused, no data output.");
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetDataList(2, new List<string>());
                DA.SetData(3, null);
                return;
            }

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
                // 1) Parser principal (siempre detecta tanto "var=?{val}unit" como "var=valor unidad")
                CalcpadSyntax.Instance.ParseVariables(content, out var names, out var values, out var units);

                // 2) Escaneo suplementario (para nombres con '.' o prime no capturados)
                SupplementalScan(content, names, values, units);

                // 3) Reordenar según orden de aparición en el archivo
                ReorderByFileAppearance(content, names, values, units);

                // 4) Crear sheet
                var sheet = new CalcpadSheet(names, values, units);
                try { sheet.SetFullCode(content); } catch { }

                if (names.Count != values.Count || values.Count != units.Count)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Mismatch in Variables/Values/Units");

                // 5) Salidas
                DA.SetDataList(0, names);
                DA.SetDataList(1, values);
                DA.SetDataList(2, units);
                DA.SetData(3, sheet);

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Loaded {names.Count} vars");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Parse error: {ex.Message}");
            }
        }

        // Escaneo suplementario: líneas simpleInput:  name = number [unit]  (sin operadores a la derecha)
        private void SupplementalScan(string content,
                                      List<string> names,
                                      List<double> values,
                                      List<string> units)
        {
            if (string.IsNullOrEmpty(content)) return;

            var existing = new HashSet<string>(names, StringComparer.Ordinal);

            foreach (Match m in RxSupplementalScan.Matches(content))
            {
                string rawName = m.Groups["name"].Value.Trim();
                if (string.IsNullOrEmpty(rawName)) continue;
                if (existing.Contains(rawName)) continue;

                if (!double.TryParse(m.Groups["num"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                    continue;

                string unit = m.Groups["unit"].Success ? m.Groups["unit"].Value : string.Empty;

                string line = GetLine(content, m.Index);
                if (RxHasRightSideOperator.IsMatch(line))
                    continue;

                names.Add(rawName);
                values.Add(val);
                units.Add(unit);
                existing.Add(rawName);
            }
        }

        // Reordenar listas según primera aparición textual (incluye asignaciones tras ';')
        private void ReorderByFileAppearance(string content,
                                             List<string> names,
                                             List<double> values,
                                             List<string> units)
        {
            if (string.IsNullOrEmpty(content) || names == null || names.Count == 0) return;

            // Map original
            var mapValue = new Dictionary<string, double>(StringComparer.Ordinal);
            var mapUnit = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int i = 0; i < names.Count; i++)
            {
                mapValue[names[i]] = i < values.Count ? values[i] : double.NaN;
                mapUnit[names[i]] = i < units.Count ? units[i] : string.Empty;
            }

            // Regex que encuentra cualquier token de variable seguido de '='
            // (también después de ';', espacios, etc.) -- ver RxAssignment (campo estático)
            var firstPos = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (Match m in RxAssignment.Matches(content))
            {
                var nm = m.Groups["name"].Value.Trim();
                if (!mapValue.ContainsKey(nm)) continue;      // sólo variables ya detectadas
                if (!firstPos.ContainsKey(nm))
                    firstPos[nm] = m.Index;                   // primera aparición
            }

            // Ordenar por posición; los no encontrados quedan al final conservando estabilidad
            var ordered = new List<string>(names);
            ordered.Sort((a, b) =>
            {
                int pa = firstPos.ContainsKey(a) ? firstPos[a] : int.MaxValue;
                int pb = firstPos.ContainsKey(b) ? firstPos[b] : int.MaxValue;
                if (pa != pb) return pa.CompareTo(pb);
                return string.CompareOrdinal(a, b);
            });

            names.Clear();
            values.Clear();
            units.Clear();
            foreach (var n in ordered)
            {
                names.Add(n);
                values.Add(mapValue.TryGetValue(n, out var v) ? v : double.NaN);
                units.Add(mapUnit.TryGetValue(n, out var u) ? u : string.Empty);
            }
        }

        private string GetLine(string text, int index)
        {
            int start = text.LastIndexOf('\n', Math.Max(0, index));
            if (start < 0) start = 0; else start++;
            int end = text.IndexOf('\n', index);
            if (end < 0) end = text.Length;
            return text.Substring(start, end - start);
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

                _watcher?.Dispose();
                _watcher = new FileSystemWatcher(dir, file)
                {
                    NotifyFilter = NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };
                _watcher.Changed += (s, e) =>
                {
                    if (_frozen) return;
                    Rhino.RhinoApp.InvokeOnUiThread(new Action(() => ExpireSolution(true)));
                };
            }
            catch { }
        }

        public override void RemovedFromDocument(Grasshopper.Kernel.GH_Document document)
        {
            try { _watcher?.Dispose(); _watcher = null; } catch { }
            base.RemovedFromDocument(document);
        }

        public override Guid ComponentGuid => new Guid("7A4CE2C1-4F7D-4C7E-A5E1-5B0C2F7E8F13");
        protected override System.Drawing.Bitmap Icon => Resources.Icon_Calc;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
