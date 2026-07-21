using System;
using System.IO;

namespace GH_Calcpad.Classes
{
    /// <summary>
    /// Exports a CalcpadSheet to HTML/PDF/DOCX by asking the worker process to render it
    /// exactly the way Calcpad's own CLI does (same worksheet template, same wkhtmltopdf.exe,
    /// same Calcpad.OpenXml writer - see Worker\GH_Calcpad.Worker\CalcpadExportWorker.cs).
    /// Replaces the previous Parser.Convert/Calcpad.Cli/Word-Interop paths entirely.
    /// </summary>
    public static class CalcpadExporter
    {
        public static (bool success, string finalPath, string error) Export(
            CalcpadSheet sheet, string outputFolder, string baseName, string format)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.OriginalCode))
                return (false, null, "No code to export. Run 'Play CPD' first.");

            if (string.IsNullOrWhiteSpace(outputFolder))
                return (false, null, "'Output Folder' not provided.");

            if (string.IsNullOrWhiteSpace(baseName))
                return (false, null, "'File' name not provided.");

            string folder = outputFolder.Trim();
            try
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
            }
            catch (Exception ex)
            {
                return (false, null, $"Could not create directory: {ex.Message}");
            }

            string safeBase = Path.GetFileNameWithoutExtension(baseName.Trim());
            string finalPath = Path.Combine(folder, safeBase + "." + format);

            try
            {
                var response = CalcpadWorkerClient.Instance.Export(sheet.OriginalCode, format, finalPath);
                return response.Ok
                    ? (true, response.FinalPath, null)
                    : (false, finalPath, response.Error);
            }
            catch (Exception ex)
            {
                return (false, finalPath, $"Calcpad worker unavailable: {ex.Message}");
            }
        }
    }
}
