using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Grasshopper.Kernel;
using GH_Calcpad.Classes;    // Namespace where we define CalcpadSheet
using GH_Calcpad.Properties;  // Resources (icon)

namespace GH_Calcpad.Components
{
    /// <summary>
    /// GH Component to read a .cpd file directly as text,
    /// extract variables, values and units without using intermediate HTML,
    /// and generate a CalcpadSheet for later use.
    /// Monitors file changes for automatic recomputation.
    /// </summary>
    public class GH_Calcpad_Load_cpd : GH_Component
    {
        private FileSystemWatcher _watcher;

        public GH_Calcpad_Load_cpd()
          : base("Load CPD", "LoadCPD",
                 "Reads a .cpd file and extracts variables, values and units",
                 "Calcpad", "2. File Loading")  // Load CPD
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("FilePath", "P", "Complete path to .cpd file", GH_ParamAccess.item);
            pManager.AddBooleanParameter("CaptureExplicit", "C", "If True, only captures variables with format variable=?{value}unit or valueunit';'variable", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Variables", "N", "Names of all variables found", GH_ParamAccess.list);
            pManager.AddNumberParameter("Values", "V", "Numeric values associated 1:1 with Variables", GH_ParamAccess.list);
            pManager.AddTextParameter("Units", "U", "Units corresponding 1:1 with Variables", GH_ParamAccess.list);
            pManager.AddGenericParameter("SheetObj", "S", "CalcpadSheet instance for later consumption", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string path = null;
            bool captureExplicit = false;
            if (!DA.GetData(0, ref path)) return;
            DA.GetData(1, ref captureExplicit);

            SetupFileWatcher(path);

            if (!File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"File does not exist:\n{path}");
                return;
            }

            try
            {
                string content = File.ReadAllText(path, Encoding.UTF8);

                var names = new List<string>();
                var values = new List<double>();
                var units = new List<string>();

                if (!captureExplicit)
                {
                    // Path 1: General functional regex
                    string pattern = @"\b([^\s=]+)\b\s*=\s*(?:([0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)|\?\s*\{\s*([0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)\s*\})([^'\r\n]*)";
                    var regex = new Regex(pattern);

                    foreach (Match m in regex.Matches(content))
                    {
                        string name = m.Groups[1].Value.Trim();
                        if (name.Contains("';'"))
                        {
                            int idx = name.LastIndexOf("';'") + 3;
                            if (idx < name.Length)
                                name = name.Substring(idx).Trim();
                        }

                        string direct = m.Groups[2].Value;
                        string braced = m.Groups[3].Value;
                        string unit = m.Groups[4].Value.Trim();

                        string num = !string.IsNullOrEmpty(direct) ? direct : braced;

                        if (string.IsNullOrEmpty(num))
                            continue;

                        if (Regex.IsMatch(num, @"[\*\^\/\+\-\(]|max|min", RegexOptions.IgnoreCase))
                            continue;

                        double val = double.NaN;
                        if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double tmp))
                            val = tmp;

                        names.Add(name);
                        values.Add(val);
                        units.Add(string.IsNullOrWhiteSpace(unit) ? string.Empty : unit);
                    }
                }
                else
                {
                    // Explicit mode: captures both formats
                    // 1. variable = ? {value}unit
                    // 2. valueunit';'variable (in middle of line)
                    string patternE1 = @"\b([^\s=]+)\b\s*=\s*\?\s*\{\s*([0-9]+(?:\.[0-9]+)?(?:[eE][+\-]?[0-9]+)?)\s*\}([^\r\n']*)";
                    string patternE2 = @"([0-9]+(?:\.[0-9]+)?(?:[eE][+\-]?[0-9]+)?)([^\r\n]*?)';'([^\s=]+)";
                    var regexE1 = new Regex(patternE1);
                    var regexE2 = new Regex(patternE2);

                    var namesSet = new HashSet<string>();

                    // 1. variable = ? {value}unit
                    foreach (Match m in regexE1.Matches(content))
                    {
                        string rawName = m.Groups[1].Value.Trim();
                        // If variable comes after a single quote
                        if (rawName.Contains("'"))
                        {
                            var parts = rawName.Split(new[] { '\'' }, StringSplitOptions.RemoveEmptyEntries);
                            rawName = parts[parts.Length - 1].Trim();
                        }
                        string numStr = m.Groups[2].Value.Trim();
                        string unit = m.Groups[3].Value.Trim();
                        if (double.TryParse(numStr, NumberStyles.Any,
                            CultureInfo.InvariantCulture, out double val))
                        {
                            if (namesSet.Add(rawName))
                            {
                                names.Add(rawName);
                                values.Add(val);
                                units.Add(unit);
                            }
                        }
                    }

                    // 2. valueunit';'variable
                    foreach (Match m in regexE2.Matches(content))
                    {
                        string numStr = m.Groups[1].Value.Trim();
                        string unit = m.Groups[2].Value.Trim();
                        string rawName = m.Groups[3].Value.Trim();
                        if (double.TryParse(numStr, NumberStyles.Any,
                            CultureInfo.InvariantCulture, out double val))
                        {
                            if (namesSet.Add(rawName))
                            {
                                names.Add(rawName);
                                values.Add(val);
                                units.Add(unit);
                            }
                        }
                    }
                }

                // ✅ Create CalcpadSheet with enhanced API
                var sheetObj = new CalcpadSheet(names, values, units);
                
                // ✅ CRITICAL: Set complete code so Calculator works
                try
                {
                    sheetObj.SetFullCode(content);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "CPD code loaded correctly into CalcpadSheet.");
                }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Error setting code in CalcpadSheet: {ex.Message}");
                }

                if (names.Count != values.Count || values.Count != units.Count)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Mismatch in Variables/Values/Units");

                DA.SetDataList(0, names);
                DA.SetDataList(1, values);
                DA.SetDataList(2, units);
                DA.SetData(3, sheetObj);

                if (captureExplicit)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Variables captured with format variable=?{value}unit and valueunit';'variable. If you want to capture all, set boolean to False.");
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error loading CPD: {ex.Message}");
            }
        }

        private void SetupFileWatcher(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
                string name = Path.GetFileName(filePath);
                if (_watcher != null)
                {
                    if (_watcher.Path == dir && _watcher.Filter == name) return;
                    _watcher.Dispose();
                }
                _watcher = new FileSystemWatcher(dir, name)
                {
                    NotifyFilter = NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };
                _watcher.Changed += (s, e) => ExpireSolution(true);
            }
            catch { }
        }

        protected override System.Drawing.Bitmap Icon => Resources.Icon_Calc;
        public override Guid ComponentGuid => new Guid("546918D1-6906-4258-9A1F-1378EA19257C");
    }
}