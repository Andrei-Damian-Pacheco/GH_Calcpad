using System;
using System.IO;
using System.Drawing;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component for saving modified CPD/TXT file after calculation.
    /// Workflow: Load → ModVar → Play → SaveCPD
    /// </summary>
    public class GH_Calcpad_Save_cpd : GH_Component
    {
        public GH_Calcpad_Save_cpd()
          : base(
                "Save CPD/TXT",
                "SaveCPD",
                "Saves modified code as .cpd (default) or .txt",
                "Calcpad",
                "6. Saving & Export"
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter(
                "Updated Sheet", "US",
                "Updated CalcpadSheet (from Play CPD)",
                GH_ParamAccess.item);

            p.AddTextParameter(
                "File", "N",
                "Base name (without extension); existing extension will be ignored",
                GH_ParamAccess.item);

            p.AddTextParameter(
                "Output Folder", "F",
                "Destination folder path",
                GH_ParamAccess.item);

            p.AddBooleanParameter(
                "CPD/TXT", "FMT",
                "False = .cpd (default), True = .txt",
                GH_ParamAccess.item, false);

            p.AddBooleanParameter(
                "Execute", "X",
                "True = save",
                GH_ParamAccess.item, false);

            // Optional (like in other exporters)
            p[1].Optional = true; // N
            p[2].Optional = true; // F
            p[3].Optional = true; // FMT
            // X not optional (for consistency with others)
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter(
                "Save Path", "SP",
                "Complete path of the saved file",
                GH_ParamAccess.item);

            p.AddBooleanParameter(
                "Success", "S",
                "True if file was saved or already up-to-date",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Inputs
            object data = null;
            string fileName = null;
            string outputFolder = null;
            bool saveAsTxt = false;
            bool execute = false;

            if (!DA.GetData(0, ref data)) return;
            DA.GetData(1, ref fileName);
            DA.GetData(2, ref outputFolder);
            DA.GetData(3, ref saveAsTxt);
            DA.GetData(4, ref execute);

            if (!execute)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Set Execute=True to save CPD/TXT");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Connect 'Output Folder'. No save until path is provided.");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Connect 'File'. No save until name is provided.");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            // 2) Unwrap
            CalcpadSheet sheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return;
            }

            // 3) Code
            string code = sheet.OriginalCode;
            if (string.IsNullOrEmpty(code))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "CalcpadSheet contains no code. Use 'Play CPD' first.");
                return;
            }

            // 4) Paths (no fallbacks)
            string baseName = Path.GetFileNameWithoutExtension(fileName.Trim());
            string folder = outputFolder.Trim();

            string ext = saveAsTxt ? ".txt" : ".cpd";
            string finalPath = Path.Combine(folder, baseName + ext);

            // 5) Ensure folder
            try
            {
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Directory created: {folder}");
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Could not create directory: {ex.Message}");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            // 6) If exists and content is equal, don't rewrite
            try
            {
                if (File.Exists(finalPath))
                {
                    string existing = File.ReadAllText(finalPath, Encoding.UTF8);
                    if (StringEqualsNormalized(existing, code))
                    {
                        DA.SetData(0, finalPath);
                        DA.SetData(1, true);
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "File already up-to-date. Not rewritten.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Skipping content check: {ex.Message}");
            }

            // 7) Safe write (tmp + replace/move)
            bool success = false;
            long fileSize = 0;
            string tmpPath = finalPath + ".tmp";

            try
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Saving ({ext}) {code.Length} chars → {baseName}{ext}");
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Target path: {finalPath}");

                File.WriteAllText(tmpPath, code, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                if (File.Exists(finalPath))
                    File.Replace(tmpPath, finalPath, null);
                else
                    File.Move(tmpPath, finalPath);

                if (File.Exists(finalPath))
                {
                    fileSize = new FileInfo(finalPath).Length;
                    success = true;
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"✅ File saved: {fileSize:N0} bytes");
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File not saved.");
                }
            }
            catch (UnauthorizedAccessException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"No permission to write to: {finalPath}");
                TryDelete(tmpPath);
            }
            catch (DirectoryNotFoundException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Directory not found: {folder}");
                TryDelete(tmpPath);
            }
            catch (PathTooLongException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Path too long: {finalPath}");
                TryDelete(tmpPath);
            }
            catch (IOException ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"I/O error: {ex.Message}");
                TryDelete(tmpPath);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error saving: {ex.Message}");
                TryDelete(tmpPath);
            }

            // 8) Outputs
            DA.SetData(0, success ? finalPath : null);
            DA.SetData(1, success);

            if (success)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"✅ Saved → {Path.GetFileName(finalPath)} | {fileSize / 1024.0:F1} KB | Variables: {sheet.Variables?.Count ?? 0}");
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "❌ Save failed");
            }
        }

        private static bool StringEqualsNormalized(string a, string b)
        {
            if (a == null || b == null) return a == b;
            string na = a.Replace("\r\n", "\n").Replace("\r", "\n");
            string nb = b.Replace("\r\n", "\n").Replace("\r", "\n");
            return string.Equals(na, nb, StringComparison.Ordinal);
        }

        private static void TryDelete(string path)
        {
            try { if (!string.IsNullOrEmpty(path) && File.Exists(path)) File.Delete(path); } catch { }
        }

        public override Guid ComponentGuid => new Guid("E5F6A7B8-C9D0-4E1F-2A3B-4C5D6E7F8A9B");
        protected override Bitmap Icon => Resources.Icon_Save;
        public override string ToString() => "GH_Calcpad_Save_cpd: CPD/TXT saver";
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}