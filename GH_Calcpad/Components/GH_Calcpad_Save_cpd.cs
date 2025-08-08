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
    /// Component for saving modified CPD file after calculation.
    /// Preserves changes made to variables and generates a new version of the file.
    /// Workflow: Load → ModVar → Play → SaveCPD
    /// </summary>
    public class GH_Calcpad_Save_cpd : GH_Component
    {
        public GH_Calcpad_Save_cpd()
          : base(
                "Save CPD",            // Component name
                "SaveCPD",             // Nickname
                "Saves modified CPD file with new variable values",
                "Calcpad",             // Category
                "6. Saving & Export"      // Save CPD
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter(
                "Updated Sheet", "US",
                "Updated CalcpadSheet (from Play CPD)",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "File Name", "N",
                "Name for the CPD file (without extension)",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "Output Folder", "F",
                "Destination folder path (optional - if empty, uses Desktop)",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Overwrite", "OW",
                "If True, overwrites existing file. If False, fails if file already exists",
                GH_ParamAccess.item, false);
            p.AddBooleanParameter(
                "Execute", "X",
                "Set to True to execute the CPD save operation",
                GH_ParamAccess.item, false);

            // Make parameters optional
            p[1].Optional = true;
            p[2].Optional = true;
            p[3].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter(
                "Save Path", "SP",
                "Complete path of the saved CPD file",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Success", "S",
                "True if file was saved successfully",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Get inputs
            object data = null;
            string fileName = null;
            string outputFolder = null;
            bool overwrite = false;
            bool execute = false;

            if (!DA.GetData(0, ref data))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Updated Sheet not received.");
                return;
            }

            DA.GetData(1, ref fileName);
            DA.GetData(2, ref outputFolder);
            DA.GetData(3, ref overwrite);
            DA.GetData(4, ref execute);

            // 2) Early exit if Execute = False
            if (!execute)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Set Execute=True to perform CPD save operation");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            // 3) Unwrap CalcpadSheet
            CalcpadSheet sheet = null;
            if (data is GH_ObjectWrapper wrapper)
            {
                sheet = wrapper.Value as CalcpadSheet;
            }
            else
            {
                sheet = data as CalcpadSheet;
            }

            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
                    "The received object is not a valid CalcpadSheet.");
                return;
            }

            // 4) Validate that sheet has CPD code available
            string cpdCode = sheet.OriginalCode;
            if (string.IsNullOrEmpty(cpdCode))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "CalcpadSheet contains no CPD code. Use a sheet from 'Load CPD' or 'Load CPDz'.");
                return;
            }

            // 5) Determine file name and paths
            string finalFileName = string.IsNullOrWhiteSpace(fileName) 
                ? $"CalcpadSave_{DateTime.Now:yyyyMMdd_HHmmss}" 
                : fileName.Trim().Replace(".cpd", "");

            string finalFolder = string.IsNullOrWhiteSpace(outputFolder) 
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) 
                : outputFolder.Trim();

            string finalPath = Path.Combine(finalFolder, finalFileName + ".cpd");

            // 6) Validate and create destination directory
            try
            {
                if (!Directory.Exists(finalFolder))
                {
                    Directory.CreateDirectory(finalFolder);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Directory created: {finalFolder}");
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Could not create directory: {ex.Message}");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            // 7) Validate existing file
            if (File.Exists(finalPath) && !overwrite)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"File already exists and Overwrite=False: {finalPath}");
                DA.SetData(0, null);
                DA.SetData(1, false);
                return;
            }

            // 8) Perform CPD file save
            bool success = false;
            long fileSize = 0;

            try
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"Saving CPD: {cpdCode.Length} chars → {finalFileName}.cpd");
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"Target path: {finalPath}");

                // Save CPD code with UTF-8 encoding
                File.WriteAllText(finalPath, cpdCode, Encoding.UTF8);

                // Verify file was saved correctly
                if (File.Exists(finalPath))
                {
                    var fileInfo = new FileInfo(finalPath);
                    fileSize = fileInfo.Length;
                    success = true;

                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                        $"✅ CPD saved successfully: {fileSize:N0} bytes");
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "CPD file was not saved correctly.");
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"No permission to write to: {finalPath}");
            }
            catch (DirectoryNotFoundException ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"Directory not found: {finalFolder}");
            }
            catch (PathTooLongException ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"Path too long: {finalPath}");
            }
            catch (IOException ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"File I/O error: {ex.Message}");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"Error saving CPD: {ex.Message}");
            }

            // 9) Set outputs
            DA.SetData(0, success ? finalPath : null);
            DA.SetData(1, success);

            // 10) Final status message
            if (success)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"✅ CPD save completed → {Path.GetFileName(finalPath)} | " +
                    $"Size: {fileSize / 1024.0:F1} KB | Variables: {sheet.Variables?.Count ?? 0}");
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "❌ CPD save failed");
            }
        }

        public override Guid ComponentGuid
            => new Guid("E5F6A7B8-C9D0-4E1F-2A3B-4C5D6E7F8A9B");

        protected override Bitmap Icon
            => Resources.Icon_Save;

        /// <summary>
        /// Additional component information
        /// </summary>
        public override string ToString()
        {
            return "GH_Calcpad_Save_cpd: Modified CPD file saver with controlled execution";
        }

        /// <summary>
        /// Exposure level in interface
        /// </summary>
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}