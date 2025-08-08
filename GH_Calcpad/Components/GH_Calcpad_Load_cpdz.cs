using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.IO.Compression;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// GH Component to read a compiled .cpdz file,
    /// unpacks the inner .cpd, extracts variables, values and units,
    /// and generates a CalcpadSheet for later use.
    /// Monitors file changes for automatic recomputation.
    /// </summary>
    public class GH_Calcpad_Load_cpdz : GH_Component
    {
        private FileSystemWatcher _watcher;

        public GH_Calcpad_Load_cpdz()
          : base("Load CPDz", "LoadCPDz",
                 "Reads a .cpdz file and extracts variables, values and units",
                 "Calcpad", "2. File Loading")
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("FilePath", "P", "Complete path to .cpdz file", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "W", "Password (not used)", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Variables", "N", "Names of variables found", GH_ParamAccess.list);
            pManager.AddNumberParameter("Values", "V", "Corresponding values", GH_ParamAccess.list);
            pManager.AddTextParameter("Units", "U", "Corresponding units", GH_ParamAccess.list);
            pManager.AddGenericParameter("SheetObj", "S", "CalcpadSheet instance", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1. Read inputs
            string path = null;
            string pwd = null;
            if (!DA.GetData(0, ref path)) return;
            DA.GetData(1, ref pwd);

            // 2. File monitor
            SetupFileWatcher(path);

            // 3. Validate .cpdz file existence
            if (!File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $".cpdz file does not exist:\n{path}");
                return;
            }

            string cpdText;
            try
            {
                // 4. Read bytes and handle Base64 → ZIP
                byte[] fileBytes = File.ReadAllBytes(path);
                byte[] zipBytes;
                try
                {
                    string fileContent = Encoding.UTF8.GetString(fileBytes);
                    zipBytes = Convert.FromBase64String(fileContent);
                }
                catch (FormatException)
                {
                    // If Base64 conversion fails, assume it's already binary ZIP
                    zipBytes = fileBytes;
                }

                // 5. Open ZIP in memory and extract the .cpd
                using (var ms = new MemoryStream(zipBytes))
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Read))
                {
                    var entry = archive.Entries
                        .FirstOrDefault(e => e.Name.EndsWith(".cpd", StringComparison.OrdinalIgnoreCase));
                    if (entry == null)
                        throw new Exception("No .cpd found inside .cpdz file.");

                    using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                        cpdText = reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error unpacking .cpdz: {ex.Message}");
                return;
            }

            // 6. Parse the text with regex from Load_cpd
            var names = new List<string>();
            var values = new List<double>();
            var units = new List<string>();

            string pattern = @"\b([^\s=]+)\b\s*=\s*(?:([0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)|\?\s*\{\s*([0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)\s*\})([^'\r\n]*)";
            var regex = new Regex(pattern);

            foreach (Match m in regex.Matches(cpdText))
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
                
                if (string.IsNullOrEmpty(num)) continue;
                if (Regex.IsMatch(num, @"[\*\^\/\+\-\(]|max|min", RegexOptions.IgnoreCase)) continue;

                // ✅ FIXED: Updated TryParse syntax for .NET Framework 4.8
                double val;
                if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out val))
                {
                    values.Add(val);
                }
                else
                {
                    values.Add(double.NaN);
                }

                names.Add(name);
                units.Add(string.IsNullOrWhiteSpace(unit) ? string.Empty : unit);
            }

            // 7. Create CalcpadSheet and set outputs
            var sheetObj = new CalcpadSheet(names, values, units);
            
            // Set complete code so Calculator works
            try
            {
                sheetObj.SetFullCode(cpdText);
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
        }

        /// <summary>
        /// Sets up a FileSystemWatcher to re-execute when file changes
        /// </summary>
        private void SetupFileWatcher(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
                string file = Path.GetFileName(filePath);
                if (_watcher != null && _watcher.Path == dir && _watcher.Filter == file)
                    return;
                    
                // ✅ FIXED: Proper disposal of existing watcher
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

        // ✅ FIXED: Changed GUID to unique value
        public override Guid ComponentGuid => new Guid("F2E1D8C7-B6A5-4F3E-8D7C-1A2B3C4D5E6F");

        protected override System.Drawing.Bitmap Icon => Resources.Icon_Form;

        public override GH_Exposure Exposure => GH_Exposure.primary;

        // ✅ FIXED: Override the existing Dispose method from GH_Component
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    if (_watcher != null)
                    {
                        _watcher.Dispose();
                        _watcher = null;
                    }
                }
                catch { }
            }
            base.Dispose(disposing);
        }
    }
}
