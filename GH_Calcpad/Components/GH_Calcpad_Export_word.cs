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
    /// Component for exporting Calcpad calculations to Word format (.docx).
    /// Uses Calcpad's native Word export functionality (equivalent to "Open With MS Word" button).
    /// </summary>
    public class GH_Calcpad_Export_word : GH_Component
    {
        public GH_Calcpad_Export_word()
          : base(
                "Export Word",          
                "ExportWord",          
                "Exports Calcpad calculations to editable Word (.docx) format using Calcpad's native export engine",
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
                "Name for the Word file (without extension)",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "Output Folder", "F",
                "Destination folder path (optional - if empty, uses Desktop)",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Execute", "X",
                "Set to True to execute the Word export operation",
                GH_ParamAccess.item, false);

            // Make parameters optional
            p[1].Optional = true;
            p[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter(
                "Word Path", "W",
                "Complete path of the generated Word file",
                GH_ParamAccess.item);
            p.AddBooleanParameter(
                "Success", "S",
                "True if Word document was generated successfully",
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
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Set Execute=True to perform Word export operation");
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
                : fileName.Trim().Replace(".docx", "");

            string finalFolder = string.IsNullOrWhiteSpace(outputFolder) 
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) 
                : outputFolder.Trim();

            string finalPath = Path.Combine(finalFolder, finalFileName + ".docx");

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

            // 7) Try Word export methods
            bool success = false;
            long fileSize = 0;
            string methodUsed = "";

            try
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                    $"Exporting to Word: {finalFileName}.docx");

                // Method 1: Use Calcpad's native Parser.Convert (equivalent to "Open With MS Word" button)
                var nativeResult = TryNativeCalcpadWordExport(originalCode, finalPath);
                if (nativeResult.success)
                {
                    success = true;
                    fileSize = nativeResult.fileSize;
                    methodUsed = "Calcpad Native Word Export";
                }
                else
                {
                    // Method 2: Use Pandoc for HTML to DOCX conversion (if available)
                    var pandocResult = TryPandocExport(htmlContent, finalPath);
                    if (pandocResult.success)
                    {
                        success = true;
                        fileSize = pandocResult.fileSize;
                        methodUsed = "Pandoc HTML→DOCX";
                    }
                    else
                    {
                        // Method 3: Create a rich HTML file that Word can open
                        var wordHtmlResult = GenerateWordCompatibleHtml(htmlContent, finalFolder, finalFileName);
                        if (wordHtmlResult.success)
                        {
                            success = true;
                            fileSize = wordHtmlResult.fileSize;
                            finalPath = wordHtmlResult.path;
                            methodUsed = "Word-Compatible HTML";
                            
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                                "Word export created HTML file that can be opened with Word. " +
                                "For true .docx format, consider installing Pandoc or ensure Calcpad native export works.");
                        }
                        else
                        {
                            // Method 4: Generate standard HTML as last resort
                            var htmlResult = GenerateStandardHtml(htmlContent, finalFolder, finalFileName);
                            if (htmlResult.success)
                            {
                                success = true;
                                fileSize = htmlResult.fileSize;
                                finalPath = htmlResult.path;
                                methodUsed = "HTML Fallback";
                                
                                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                                    "Word export failed. HTML file created instead. " +
                                    "Install Pandoc or ensure Calcpad supports Word export.");
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
                    "❌ Export failed. Consider installing Pandoc for HTML to DOCX conversion.");
            }
        }

        /// <summary>
        /// Method 1: Use Calcpad's native Word export functionality (equivalent to "Open With MS Word" button)
        /// </summary>
        private (bool success, long fileSize) TryNativeCalcpadWordExport(string originalCode, string docxPath)
        {
            try
            {
                if (string.IsNullOrEmpty(originalCode))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "No original code available for native Word export");
                    return (false, 0);
                }

                // Create temporary calcpad file
                string tempCalcpadPath = Path.Combine(Path.GetTempPath(), $"calcpad_temp_{DateTime.Now.Ticks}.cpd");
                File.WriteAllText(tempCalcpadPath, originalCode, System.Text.Encoding.UTF8);
                
                // Use Parser.Convert - this should work for Word if PyCalcpad supports it
                var parser = new Parser();
                parser.Settings = new Settings 
                { 
                    Math = new MathSettings { Decimals = 15 }
                };

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Attempting native Calcpad Word export...");

                // Try converting directly to .docx
                bool convertResult = parser.Convert(tempCalcpadPath, docxPath);
                
                // Cleanup temp file
                try { File.Delete(tempCalcpadPath); } catch { }

                if (convertResult && File.Exists(docxPath))
                {
                    var fileInfo = new FileInfo(docxPath);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        $"✅ Word document generated with Calcpad native export: {fileInfo.Length:N0} bytes");
                    return (true, fileInfo.Length);
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        "Calcpad native Word export failed (Word export may not be available in PyCalcpad)");
                    return (false, 0);
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Native Calcpad Word export failed: {ex.Message}");
                return (false, 0);
            }
        }

        /// <summary>
        /// Method 2: Use Pandoc for HTML to DOCX conversion (professional document converter)
        /// </summary>
        private (bool success, long fileSize) TryPandocExport(string htmlContent, string docxPath)
        {
            try
            {
                // Look for Pandoc in common locations
                string[] pandocPaths = {
                    @"C:\Program Files\Pandoc\pandoc.exe",
                    @"C:\Program Files (x86)\Pandoc\pandoc.exe",
                    @"pandoc.exe" // If in PATH
                };

                string pandocPath = null;
                foreach (var path in pandocPaths)
                {
                    if (File.Exists(path) || path == "pandoc.exe")
                    {
                        try
                        {
                            // Test if pandoc is accessible
                            using (var process = Process.Start(new ProcessStartInfo
                            {
                                FileName = path,
                                Arguments = "--version",
                                UseShellExecute = false,
                                CreateNoWindow = true,
                                RedirectStandardOutput = true
                            }))
                            {
                                process.WaitForExit(5000);
                                if (process.ExitCode == 0)
                                {
                                    pandocPath = path;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                if (pandocPath == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        "Pandoc not found. Download from: https://pandoc.org/installing.html");
                    return (false, 0);
                }

                // Create temporary HTML file with proper structure
                string tempHtml = Path.Combine(Path.GetTempPath(), $"calcpad_{DateTime.Now.Ticks}.html");
                string styledHtml = CreateWordCompatibleHtmlContent(htmlContent);
                File.WriteAllText(tempHtml, styledHtml, System.Text.Encoding.UTF8);

                // Use Pandoc to convert HTML to DOCX
                var startInfo = new ProcessStartInfo
                {
                    FileName = pandocPath,
                    Arguments = $"-f html -t docx -o \"{docxPath}\" \"{tempHtml}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Converting HTML to DOCX with Pandoc...");

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit(30000); // 30 second timeout
                    
                    if (process.ExitCode == 0 && File.Exists(docxPath))
                    {
                        var fileInfo = new FileInfo(docxPath);
                        
                        // Cleanup
                        try { File.Delete(tempHtml); } catch { }
                        
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                            $"✅ Word document generated with Pandoc: {fileInfo.Length:N0} bytes");
                        return (true, fileInfo.Length);
                    }
                    else
                    {
                        string error = process.StandardError.ReadToEnd();
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                            $"Pandoc failed with exit code {process.ExitCode}: {error}");
                    }
                }

                // Cleanup on failure
                try { File.Delete(tempHtml); } catch { }
                return (false, 0);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Pandoc export failed: {ex.Message}");
                return (false, 0);
            }
        }

        /// <summary>
        /// Method 3: Generate Word-compatible HTML file
        /// </summary>
        private (bool success, long fileSize, string path) GenerateWordCompatibleHtml(string htmlContent, string folder, string fileName)
        {
            try
            {
                string htmlPath = Path.Combine(folder, fileName + ".html");
                string wordHtml = CreateWordCompatibleHtmlContent(htmlContent);
                
                File.WriteAllText(htmlPath, wordHtml, System.Text.Encoding.UTF8);
                
                if (File.Exists(htmlPath))
                {
                    var fileInfo = new FileInfo(htmlPath);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                        $"✅ Word-compatible HTML file generated: {fileInfo.Length:N0} bytes");
                    return (true, fileInfo.Length, htmlPath);
                }
                
                return (false, 0, "");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Word-compatible HTML generation failed: {ex.Message}");
                return (false, 0, "");
            }
        }

        /// <summary>
        /// Method 4: Generate standard HTML file as fallback
        /// </summary>
        private (bool success, long fileSize, string path) GenerateStandardHtml(string htmlContent, string folder, string fileName)
        {
            try
            {
                string htmlPath = Path.Combine(folder, fileName + ".html");
                string styledHtml = CreateStandardHtmlContent(htmlContent);
                
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
        /// Creates HTML optimized for Word import with proper MSO styles
        /// </summary>
        private string CreateWordCompatibleHtmlContent(string htmlContent)
        {
            return $@"<!DOCTYPE html>
<html xmlns:o=""urn:schemas-microsoft-com:office:office""
      xmlns:w=""urn:schemas-microsoft-com:office:word""
      xmlns=""http://www.w3.org/TR/REC-html40"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""Generator"" content=""Calcpad Grasshopper Plugin"">
    <meta name=""ProgId"" content=""Word.Document"">
    <meta name=""Originator"" content=""Microsoft Word"">
    <title>Calcpad Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}</title>
    
    <!--[if gte mso 9]>
    <xml>
        <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotPromptForConvert/>
            <w:DoNotShowRevisions/>
            <w:DoNotPrintRevisions/>
            <w:DisplayHorizontalDrawingGridEvery>0</w:DisplayHorizontalDrawingGridEvery>
            <w:DisplayVerticalDrawingGridEvery>2</w:DisplayVerticalDrawingGridEvery>
            <w:UseMarginsForDrawingGridOrigin/>
            <w:ValidateAgainstSchemas/>
            <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
            <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
            <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
        </w:WordDocument>
    </xml>
    <![endif]-->
    
    <style>
        /* Microsoft Word compatible styles */
        @page {{            
            size: 8.5in 11.0in;
            margin: 1.0in 1.25in 1.0in 1.25in;
            mso-header-margin: 0.5in;
            mso-footer-margin: 0.5in;
            mso-paper-source: 0;
        }}
        
        body {{            
            font-family: 'Times New Roman', serif;
            font-size: 12.0pt;
            line-height: 115%;
            margin: 0;
            padding: 0;
            color: black;
            background: white;
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 24pt;
            border-bottom: 1pt solid #333333;
            padding-bottom: 12pt;
            page-break-after: avoid;
        }}
        
        .header h1 {{
            font-size: 18.0pt;
            font-weight: bold;
            margin: 0 0 6pt 0;
            color: #333333;
        }}
        
        .header p {{
            font-size: 10.0pt;
            margin: 3pt 0;
            color: #666666;
        }}
        
        /* Calcpad equation styling for Word */
        .eq {{
            margin: 6pt 0;
            padding: 3pt 0;
            font-size: 12.0pt;
            line-height: 115%;
            font-family: 'Times New Roman', serif;
        }}
        
        /* Variable styling */
        var {{
            font-style: italic;
            font-weight: bold;
            color: black;
        }}
        
        /* Result values */
        .result {{
            color: #0066CC;
            font-weight: bold;
        }}
        
        /* Headers compatible with Word styles */
        h1 {{
            font-size: 16.0pt;
            font-weight: bold;
            margin: 18pt 0 6pt 0;
            color: #333333;
            border-bottom: 0.5pt solid #CCCCCC;
            padding-bottom: 3pt;
            page-break-after: avoid;
        }}
        
        h2 {{
            font-size: 14.0pt;
            font-weight: bold;
            margin: 12pt 0 6pt 0;
            color: #333333;
            page-break-after: avoid;
        }}
        
        h3 {{
            font-size: 13.0pt;
            font-weight: bold;
            margin: 12pt 0 6pt 0;
            color: #333333;
            page-break-after: avoid;
        }}
        
        /* Tables compatible with Word */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 6pt 0;
            font-size: 11.0pt;
        }}
        
        td, th {{
            border: 0.5pt solid #CCCCCC;
            padding: 4pt 6pt;
            text-align: left;
            vertical-align: top;
        }}
        
        th {{
            background-color: #F5F5F5;
            font-weight: bold;
        }}
        
        /* Page breaks */
        .page-break {{
            page-break-before: always;
        }}
        
        /* Math and formula styling */
        .math {{
            font-family: 'Times New Roman', serif;
            font-style: normal;
        }}
        
        /* MSO specific styles */
        .MsoNormal {{
            margin: 0;
            font-size: 12.0pt;
            font-family: 'Times New Roman', serif;
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

        /// <summary>
        /// Creates standard HTML document
        /// </summary>
        private string CreateStandardHtmlContent(string htmlContent)
        {
            return $@"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Calcpad Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}</title>
    <style>
        body {{
            font-family: 'Times New Roman', serif;
            line-height: 1.4;
            margin: 20px;
            color: #000;
            background: #fff;
            font-size: 12pt;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #333;
            padding-bottom: 15px;
        }}
        .header h1 {{
            margin: 0 0 10px 0;
            font-size: 20pt;
            font-weight: bold;
        }}
        .header p {{
            margin: 5px 0;
            font-size: 10pt;
            color: #666;
        }}
        .eq {{
            margin: 8px 0;
            padding: 4px 0;
            font-size: 12pt;
        }}
        var {{
            font-style: italic;
            font-weight: bold;
        }}
        .result {{
            color: #0066cc;
            font-weight: bold;
        }}
    </style>
</head>
<body>
    <div class=""header"">
        <h1>Calcpad Calculation Report</h1>
        <p>Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>
        <p>Exported from Grasshopper Calcpad Plugin</p>
    </div>
    <div class=""content"">
        {htmlContent}
    </div>
</body>
</html>";
        }

        public override Guid ComponentGuid
            => new Guid("C3F2D4E5-F6A7-4B8C-9D0E-1F2A3B4C5D6F");

        protected override Bitmap Icon
            => Resources.Icon_Word;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
