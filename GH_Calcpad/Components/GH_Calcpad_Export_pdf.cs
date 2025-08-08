using System;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;
using PyCalcpad;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component for exporting Calcpad calculations to PDF format.
    /// Uses Calcpad's native PDF export functionality (equivalent to "Save as PDF" button).
    /// </summary>
    public class GH_Calcpad_Export_pdf : GH_Component
    {
        public GH_Calcpad_Export_pdf()
          : base(
                "Export PDF",           
                "ExportPDF",           
                "Exports Calcpad calculations to PDF format using Calcpad's native export engine",
                "Calcpad",             
                "6. Saving & Export"   
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter(
                "Updated Sheet", "US",
                "Calculated CalcpadSheet (from Play CPD UpdatedSheet output)",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "File Name", "N",
                "Name for the PDF file (without extension)",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "Output Folder", "F",
                "Destination folder path (optional - if empty, uses Desktop)",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Execute", "X",
                "Set to True to execute the PDF export operation",
                GH_ParamAccess.item, false);

            // Make parameters optional
            p[1].Optional = true;
            p[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter(
                "PDF Path", "P",
                "Complete path of the generated PDF file",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Success", "S",
                "True if PDF was generated successfully",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Get inputs
            object data = null;
            string fileName = null;
            string outputFolder = null;
            bool execute = false;

            if (!DA.GetData(0, ref data))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Updated Sheet not received.");
                return;
            }

            DA.GetData(1, ref fileName);
            DA.GetData(2, ref outputFolder);
            DA.GetData(3, ref execute);

            // 2) Early exit if Execute = False
            if (!execute)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Set Execute=True to perform PDF export operation");
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

            // 4) Validate that sheet has content
            string htmlContent = sheet.LastHtmlResult;
            string originalCode = sheet.OriginalCode;
            
            if (string.IsNullOrEmpty(htmlContent) && string.IsNullOrEmpty(originalCode))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "CalcpadSheet contains no content. Execute 'Play CPD' with Success=True first.");
                return;
            }

            // 5) Determine file name and paths
            string finalFileName = string.IsNullOrWhiteSpace(fileName) 
                ? $"CalcpadExport_{DateTime.Now:yyyyMMdd_HHmmss}" 
                : fileName.Trim().Replace(".pdf", "");

            string finalFolder = string.IsNullOrWhiteSpace(outputFolder) 
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) 
                : outputFolder.Trim();

            string finalPath = Path.Combine(finalFolder, finalFileName + ".pdf");

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

            // 7) Try native PDF export methods
            bool success = false;
            long fileSize = 0;
            string methodUsed = "";

            try
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"Exporting to PDF: {finalFileName}.pdf");

                // Method 1: Use Calcpad's native Parser.Convert (equivalent to "Save as PDF" button)
                var nativeResult = TryNativeCalcpadExport(originalCode, finalPath);
                if (nativeResult.success)
                {
                    success = true;
                    fileSize = nativeResult.fileSize;
                    methodUsed = "Calcpad Native Export";
                }
                else
                {
                    // Method 2: Use wkhtmltopdf directly (what Calcpad uses internally)
                    var wkhtmlResult = TryWkhtmltopdfExport(htmlContent, finalPath);
                    if (wkhtmlResult.success)
                    {
                        success = true;
                        fileSize = wkhtmlResult.fileSize;
                        methodUsed = "wkhtmltopdf Direct";
                    }
                    else
                    {
                        // Method 3: Use browser PDF printing as fallback
                        var browserResult = TryBrowserPdfExport(htmlContent, finalPath);
                        if (browserResult.success)
                        {
                            success = true;
                            fileSize = browserResult.fileSize;
                            methodUsed = "Browser PDF";
                        }
                        else
                        {
                            // Method 4: Generate HTML as last resort
                            var htmlResult = GenerateHtmlFile(htmlContent, finalFolder, finalFileName);
                            if (htmlResult.success)
                            {
                                success = true;
                                fileSize = htmlResult.fileSize;
                                finalPath = htmlResult.path;
                                methodUsed = "HTML Fallback";
                                
                                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                                    "PDF generation failed. HTML file created instead. " +
                                    "Install wkhtmltopdf or ensure Chrome/Edge is available for PDF export.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Export error: {ex.Message}");
                methodUsed = "Error";
            }

            // 8) Set outputs
            DA.SetData(0, success ? finalPath : null);
            DA.SetData(1, success);

            // 9) Final status message
            if (success)
            {
                string extension = Path.GetExtension(finalPath).ToUpper();
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"✅ {extension} export completed → {Path.GetFileName(finalPath)} | " +
                    $"Size: {fileSize / 1024.0:F1} KB | Method: {methodUsed}");
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
                    "❌ Export failed. Consider installing wkhtmltopdf from https://wkhtmltopdf.org/");
            }
        }

        /// <summary>
        /// Method 1: Use Calcpad's native export functionality (equivalent to "Save as PDF" button)
        /// </summary>
        private (bool success, long fileSize) TryNativeCalcpadExport(string originalCode, string pdfPath)
        {
            try
            {
                if (string.IsNullOrEmpty(originalCode))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "No original code available for native export");
                    return (false, 0);
                }

                // Create temporary calcpad file
                string tempCalcpadPath = Path.Combine(Path.GetTempPath(), $"calcpad_temp_{DateTime.Now.Ticks}.cpd");
                File.WriteAllText(tempCalcpadPath, originalCode, System.Text.Encoding.UTF8);
                
                // Use Parser.Convert - this is what Calcpad uses internally for "Save as PDF"
                var parser = new Parser();
                parser.Settings = new Settings 
                { 
                    Math = new MathSettings { Decimals = 15 }
                };

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Attempting native Calcpad PDF export...");

                // This should work if PyCalcpad includes wkhtmltopdf dependencies
                bool convertResult = parser.Convert(tempCalcpadPath, pdfPath);
                
                // Cleanup temp file
                try { File.Delete(tempCalcpadPath); } catch { }

                if (convertResult && File.Exists(pdfPath))
                {
                    var fileInfo = new FileInfo(pdfPath);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        $"✅ PDF generated with Calcpad native export: {fileInfo.Length:N0} bytes");
                    return (true, fileInfo.Length);
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        "Calcpad native export failed (wkhtmltopdf may not be available)");
                    return (false, 0);
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Native Calcpad export failed: {ex.Message}");
                return (false, 0);
            }
        }

        /// <summary>
        /// Method 2: Use wkhtmltopdf directly (same tool Calcpad uses)
        /// </summary>
        private (bool success, long fileSize) TryWkhtmltopdfExport(string htmlContent, string pdfPath)
        {
            try
            {
                // Look for wkhtmltopdf in common locations
                string[] wkhtmlPaths = {
                    @"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
                    @"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
                    // Check if it's bundled with Calcpad
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Calcpad", "wkhtmltopdf.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Calcpad", "wkhtmltopdf.exe")
                };

                string wkhtmlPath = null;
                foreach (var path in wkhtmlPaths)
                {
                    if (File.Exists(path))
                    {
                        wkhtmlPath = path;
                        break;
                    }
                }

                if (wkhtmlPath == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        "wkhtmltopdf not found. Download from: https://wkhtmltopdf.org/downloads.html");
                    return (false, 0);
                }

                // Create temporary HTML file with proper styling
                string tempHtml = Path.Combine(Path.GetTempPath(), $"calcpad_{DateTime.Now.Ticks}.html");
                string styledHtml = CreateStandaloneHtml(htmlContent);
                File.WriteAllText(tempHtml, styledHtml, System.Text.Encoding.UTF8);

                // Configure wkhtmltopdf for professional output
                var startInfo = new ProcessStartInfo
                {
                    FileName = wkhtmlPath,
                    Arguments = $"--page-size A4 --orientation Portrait " +
                               $"--margin-top 20mm --margin-right 15mm --margin-bottom 20mm --margin-left 15mm " +
                               $"--encoding UTF-8 --disable-smart-shrinking " +
                               $"\"{tempHtml}\" \"{pdfPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Generating PDF with wkhtmltopdf...");

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit(30000); // 30 second timeout
                    
                    if (process.ExitCode == 0 && File.Exists(pdfPath))
                    {
                        var fileInfo = new FileInfo(pdfPath);
                        
                        // Cleanup
                        try { File.Delete(tempHtml); } catch { }
                        
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                            $"✅ PDF generated with wkhtmltopdf: {fileInfo.Length:N0} bytes");
                        return (true, fileInfo.Length);
                    }
                    else
                    {
                        string error = process.StandardError.ReadToEnd();
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                            $"wkhtmltopdf failed with exit code {process.ExitCode}: {error}");
                    }
                }

                // Cleanup on failure
                try { File.Delete(tempHtml); } catch { }
                return (false, 0);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"wkhtmltopdf export failed: {ex.Message}");
                return (false, 0);
            }
        }

        /// <summary>
        /// Method 3: Use browser PDF printing (Chrome/Edge)
        /// </summary>
        private (bool success, long fileSize) TryBrowserPdfExport(string htmlContent, string pdfPath)
        {
            try
            {
                // Look for Chrome or Edge
                string[] browserPaths = {
                    @"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                    @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
                };

                string browserPath = null;
                foreach (var path in browserPaths)
                {
                    if (File.Exists(path))
                    {
                        browserPath = path;
                        break;
                    }
                }

                if (browserPath == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Chrome/Edge not found for PDF printing");
                    return (false, 0);
                }

                // Create temporary HTML file
                string tempHtml = Path.Combine(Path.GetTempPath(), $"calcpad_{DateTime.Now.Ticks}.html");
                string styledHtml = CreateStandaloneHtml(htmlContent);
                File.WriteAllText(tempHtml, styledHtml, System.Text.Encoding.UTF8);

                // Use browser headless PDF printing
                var startInfo = new ProcessStartInfo
                {
                    FileName = browserPath,
                    Arguments = $"--headless --disable-gpu --run-all-compositor-stages-before-draw " +
                               $"--print-to-pdf=\"{pdfPath}\" --print-to-pdf-no-header " +
                               $"--virtual-time-budget=2000 \"file:///{tempHtml.Replace('\\', '/')}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Generating PDF with browser...");

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit(30000); // 30 second timeout
                    
                    // Give some time for file to be written
                    System.Threading.Thread.Sleep(1000);
                    
                    if (File.Exists(pdfPath))
                    {
                        var fileInfo = new FileInfo(pdfPath);
                        
                        // Cleanup
                        try { File.Delete(tempHtml); } catch { }
                        
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                            $"✅ PDF generated with browser: {fileInfo.Length:N0} bytes");
                        return (true, fileInfo.Length);
                    }
                }

                // Cleanup on failure
                try { File.Delete(tempHtml); } catch { }
                return (false, 0);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Browser PDF export failed: {ex.Message}");
                return (false, 0);
            }
        }

        /// <summary>
        /// Method 4: Generate high-quality HTML file as fallback
        /// </summary>
        private (bool success, long fileSize, string path) GenerateHtmlFile(string htmlContent, string folder, string fileName)
        {
            try
            {
                string htmlPath = Path.Combine(folder, fileName + ".html");
                string styledHtml = CreateStandaloneHtml(htmlContent);
                
                File.WriteAllText(htmlPath, styledHtml, System.Text.Encoding.UTF8);
                
                if (File.Exists(htmlPath))
                {
                    var fileInfo = new FileInfo(htmlPath);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        $"✅ HTML file generated: {fileInfo.Length:N0} bytes");
                    return (true, fileInfo.Length, htmlPath);
                }
                
                return (false, 0, "");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"HTML generation failed: {ex.Message}");
                return (false, 0, "");
            }
        }

        /// <summary>
        /// Creates a professional, print-ready HTML document (compatible with Calcpad styling)
        /// </summary>
        private string CreateStandaloneHtml(string htmlContent)
        {
            return $@"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Calcpad Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}</title>
    <style>
        @page {{
            size: A4;
            margin: 20mm 15mm;
        }}
        
        body {{
            font-family: 'Times New Roman', serif;
            line-height: 1.4;
            margin: 0;
            padding: 0;
            color: #000;
            background: #fff;
            font-size: 12pt;
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #333;
            padding-bottom: 15px;
            page-break-after: avoid;
        }}
        
        .header h1 {{
            margin: 0 0 10px 0;
            font-size: 20pt;
            font-weight: bold;
            color: #333;
        }}
        
        .header p {{
            margin: 5px 0;
            font-size: 10pt;
            color: #666;
        }}
        
        /* Calcpad equation styling */
        .eq {{
            margin: 8px 0;
            padding: 4px 0;
            font-size: 12pt;
            line-height: 1.3;
        }}
        
        /* Variable styling */
        var {{
            font-style: italic;
            font-weight: bold;
            color: #000;
        }}
        
        /* Result values */
        .result {{
            color: #0066cc;
            font-weight: bold;
        }}
        
        /* Headers */
        h1, h2, h3, h4, h5, h6 {{
            color: #333;
            border-bottom: 1px solid #ddd;
            padding-bottom: 3px;
            margin-top: 20px;
            margin-bottom: 10px;
            page-break-after: avoid;
        }}
        
        h1 {{ font-size: 16pt; }}
        h2 {{ font-size: 14pt; }}
        h3 {{ font-size: 13pt; }}
        
        /* Tables */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0;
            font-size: 11pt;
        }}
        
        td, th {{
            border: 1px solid #ddd;
            padding: 6px;
            text-align: left;
        }}
        
        th {{
            background-color: #f5f5f5;
            font-weight: bold;
        }}
        
        /* Page breaks */
        .page-break {{
            page-break-before: always;
        }}
        
        /* Print optimizations */
        @media print {{
            body {{ 
                margin: 0; 
                -webkit-print-color-adjust: exact;
                color-adjust: exact;
            }}
            .no-print {{ display: none; }}
            .eq {{ break-inside: avoid; }}
            h1, h2, h3, h4, h5, h6 {{ break-after: avoid; }}
        }}
        
        /* Calcpad-specific styling */
        .calcpad-content {{
            font-family: 'Times New Roman', serif;
        }}
        
        .math {{
            font-family: 'Times New Roman', serif;
            font-style: normal;
        }}
    </style>
</head>
<body>
    <div class=""header"">
        <h1>Calcpad Calculation Report</h1>
        <p>Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>
        <p>Exported from Grasshopper Calcpad Plugin</p>
    </div>
    <div class=""calcpad-content"">
        {htmlContent}
    </div>
</body>
</html>";
        }

        public override Guid ComponentGuid
            => new Guid("B2F1C4D5-E6A7-4B8C-9D0E-1F2A3B4C5D6E");

        protected override Bitmap Icon
            => Resources.Icon_Pdf;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}