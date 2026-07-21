using Calcpad.Core;
using Calcpad.OpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using GH_Calcpad.Protocol;

namespace GH_Calcpad.Worker
{
    // Mirrors exactly what Calcpad.Cli's own TryConvertOnStartup/Converter do (see the
    // real Calcpad source): parse with getXml only for .docx, wrap the rendered HTML with
    // Calcpad's own worksheet template (same CSS the desktop app and CLI use), then either
    // write the HTML directly, hand it to Calcpad.OpenXml's OpenXmlWriter for a native
    // .docx, or shell out to the same wkhtmltopdf.exe Calcpad.Cli uses for .pdf.
    internal static class CalcpadExportWorker
    {
        private static readonly string AssetsDir = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppContext.BaseDirectory,
            "Assets");

        private static readonly Lazy<string> HtmlWorksheetTemplate = new(() =>
            File.ReadAllText(Path.Combine(AssetsDir, "template.html")));

        public static SolveResponse Export(string id, string code, string format, string outputPath)
        {
            var response = new SolveResponse { Id = id };
            try
            {
                if (string.IsNullOrWhiteSpace(outputPath))
                    throw new ArgumentException("Export request must include 'outputPath'.");

                bool isDocx = string.Equals(format, "docx", StringComparison.OrdinalIgnoreCase);
                var settings = new Settings();
                settings.Math.Decimals = 15;
                var parser = new ExpressionParser { Settings = settings };
                parser.Parse(code, true, isDocx);

                var fullHtml = HtmlWorksheetTemplate.Value + parser.HtmlResult + " </body></html>";

                switch (format?.ToLowerInvariant())
                {
                    case "html":
                        File.WriteAllText(outputPath, fullHtml);
                        break;
                    case "docx":
                        new OpenXmlWriter(parser.OpenXmlExpressions).Convert(fullHtml, outputPath);
                        break;
                    case "pdf":
                        RenderPdf(fullHtml, outputPath);
                        break;
                    default:
                        throw new ArgumentException($"Unknown export format '{format}'. Expected html, pdf or docx.");
                }

                response.Ok = true;
                response.FinalPath = outputPath;
            }
            catch (Exception ex)
            {
                response.Ok = false;
                response.Error = ex.Message;
            }
            return response;
        }

        private static void RenderPdf(string html, string outputPath)
        {
            var tmpHtml = Path.Combine(Path.GetTempPath(), $"calcpad_export_{Guid.NewGuid():N}.html");
            File.WriteAllText(tmpHtml, html);
            try
            {
                var wkhtmltopdf = Path.Combine(AssetsDir, "wkhtmltopdf.exe");
                if (!File.Exists(wkhtmltopdf))
                    throw new FileNotFoundException("wkhtmltopdf.exe not found next to the worker.", wkhtmltopdf);

                // Same arguments Calcpad.Cli's own Converter.ToPdf uses.
                var psi = new ProcessStartInfo
                {
                    FileName = wkhtmltopdf,
                    Arguments = "--enable-local-file-access --disable-smart-shrinking --page-size A4 " +
                                "--margin-bottom 15 --margin-left 15 --margin-right 10 --margin-top 15 " +
                                $"\"{tmpHtml}\" \"{outputPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true
                };

                using var process = Process.Start(psi);
                if (!process.WaitForExit(60000))
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("wkhtmltopdf did not finish within 60 seconds.");
                }

                if (!File.Exists(outputPath))
                {
                    var stderr = process.StandardError.ReadToEnd();
                    throw new InvalidOperationException(
                        $"wkhtmltopdf did not produce a PDF.{(string.IsNullOrWhiteSpace(stderr) ? "" : $" Details: {stderr}")}");
                }
            }
            finally
            {
                try { File.Delete(tmpHtml); } catch { }
            }
        }
    }
}
