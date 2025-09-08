using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Diagnostics; // Added
using PyCalcpad;

namespace GH_Calcpad.Classes
{
    public enum ExportFormat { Pdf, Docx, Html }

    public static class CalcpadExporter
    {
        // Resolver to ensure DocumentFormat.OpenXml 3.3.0 is loaded from the plugin folder
        static CalcpadExporter()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    var name = new AssemblyName(args.Name);
                    // Intervene only for OpenXml (and potentially other dependencies if needed)
                    if (!name.Name.Equals("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase))
                        return null;

                    var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    var candidate = Path.Combine(baseDir, "DocumentFormat.OpenXml.dll");
                    if (File.Exists(candidate))
                        return Assembly.LoadFrom(candidate);

                    // Additional attempt: same folder as current assembly
                    var here = Path.GetDirectoryName(typeof(CalcpadExporter).Assembly.Location);
                    if (!string.IsNullOrEmpty(here))
                    {
                        candidate = Path.Combine(here, "DocumentFormat.OpenXml.dll");
                        if (File.Exists(candidate))
                            return Assembly.LoadFrom(candidate);
                    }
                }
                catch { /* ignore */ }

                return null; // allow normal resolution
            };
        }

        // ========== Unified public API ==========

        public static (bool success, string finalPath, long size, string method) ExportPdfNative(
            CalcpadSheet sheet, string outputFolder, string baseName, Action<string> log = null)
            => ExportViaConvert(sheet, outputFolder, baseName, ".pdf", log);

        public static (bool success, string finalPath, long size, string method) ExportHtmlNative(
            CalcpadSheet sheet, string outputFolder, string baseName, Action<string> log = null)
            => ExportViaConvert(sheet, outputFolder, baseName, ".html", log);

        // Native DOCX (if Convert supports .docx in this build)
        public static (bool success, string finalPath, long size, string method) ExportDocxNative(
            CalcpadSheet sheet, string outputFolder, string baseName, Action<string> log = null)
            => ExportViaConvert(sheet, outputFolder, baseName, ".docx", log);

        // DOCX via Word Interop (late-binding) + OMath
        public static (bool success, string finalPath, long size, string method) ExportDocxInterop(
            CalcpadSheet sheet, string outputFolder, string baseName, Action<string> log = null)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.OriginalCode))
                return (false, null, 0, "NoCode");

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                log?.Invoke("'Output Folder' not provided. Export cancelled.");
                return (false, null, 0, "NoOutputFolder");
            }
            if (string.IsNullOrWhiteSpace(baseName))
            {
                log?.Invoke("'File' name not provided. Export cancelled.");
                return (false, null, 0, "NoFileName");
            }

            string folder = outputFolder.Trim();

            try
            {
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                    log?.Invoke($"Directory created: {folder}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"Could not create directory: {ex.Message}");
                return (false, null, 0, "DirError");
            }

            string safeBase = Path.GetFileNameWithoutExtension(baseName.Trim());
            string docxPath = Path.Combine(folder, safeBase + ".docx");

            var res = BuildDocxWithOMathSTA(sheet, docxPath, log);
            return res.success
                ? (true, docxPath, res.size, "Word Interop OMath (STA)")
                : (false, docxPath, 0, "InteropFailed");
        }

        // DOCX via Calcpad CLI (native OpenXml out-of-process)
        public static (bool success, string finalPath, long size, string method) ExportDocxCli(
            CalcpadSheet sheet, string outputFolder, string baseName, string cliPath = null, Action<string> log = null)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.OriginalCode))
                return (false, null, 0, "NoCode");

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                log?.Invoke("'Output Folder' not provided. Export cancelled.");
                return (false, null, 0, "NoOutputFolder");
            }
            if (string.IsNullOrWhiteSpace(baseName))
            {
                log?.Invoke("'File' name not provided. Export cancelled.");
                return (false, null, 0, "NoFileName");
            }

            string folder = outputFolder.Trim();
            try
            {
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                    log?.Invoke($"Directory created: {folder}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"Could not create directory: {ex.Message}");
                return (false, null, 0, "DirError");
            }

            string safeBase = Path.GetFileNameWithoutExtension(baseName.Trim());
            string docxPath = Path.Combine(folder, safeBase + ".docx");

            // Write temporary CPD
            string tmpCpd = Path.Combine(Path.GetTempPath(), $"calcpad_{DateTime.Now.Ticks}.cpd");
            try { File.WriteAllText(tmpCpd, sheet.OriginalCode, new UTF8Encoding(false)); }
            catch (Exception ex)
            {
                log?.Invoke($"Could not write temporary CPD: {ex.Message}");
                return (false, docxPath, 0, "TmpWriteFailed");
            }

            // Locate CLI
            string exe = ResolveCalcpadCliPath(cliPath, log);
            if (string.IsNullOrEmpty(exe) || !File.Exists(exe))
            {
                log?.Invoke("Calcpad.Cli not found. Define CALCPAD_CLI environment variable or provide a path.");
                TryDelete(tmpCpd);
                return (false, docxPath, 0, "CliNotFound");
            }

            // Try several common argument patterns
            var candidates = new[]
            {
                // 1) simple style: Calcpad.Cli.exe "in.cpd" "out.docx"
                $"\"{tmpCpd}\" \"{docxPath}\"",
                // 2) with command: Calcpad.Cli.exe convert "in.cpd" "out.docx"
                $"convert \"{tmpCpd}\" \"{docxPath}\"",
                // 3) with flags
                $"-i \"{tmpCpd}\" -o \"{docxPath}\" -f docx"
            };

            bool ok = false;
            string lastErr = null;

            foreach (var args in candidates)
            {
                log?.Invoke($"Running CLI: {Path.GetFileName(exe)} {args}");
                var (runOk, err) = RunProcess(exe, args, log, timeoutMs: 60000);
                lastErr = err;
                if (runOk && File.Exists(docxPath))
                {
                    ok = true;
                    break;
                }
            }

            TryDelete(tmpCpd);

            if (ok)
            {
                long size = new FileInfo(docxPath).Length;
                return (true, docxPath, size, "Calcpad CLI");
            }

            log?.Invoke($"CLI did not generate DOCX. {(string.IsNullOrWhiteSpace(lastErr) ? "" : $"Details: {lastErr}")}");
            return (false, docxPath, 0, "CliFailed");
        }

        // ========== Internal implementations ==========

        private static (bool success, string finalPath, long size, string method) ExportViaConvert(
            CalcpadSheet sheet, string outputFolder, string baseName, string extension, Action<string> log)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.OriginalCode))
                return (false, null, 0, "NoCode");

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                log?.Invoke("'Output Folder' not provided. Export cancelled.");
                return (false, null, 0, "NoOutputFolder");
            }
            if (string.IsNullOrWhiteSpace(baseName))
            {
                log?.Invoke("'File' name not provided. Export cancelled.");
                return (false, null, 0, "NoFileName");
            }

            string folder = outputFolder.Trim();

            try
            {
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                    log?.Invoke($"Directory created: {folder}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"Could not create directory: {ex.Message}");
                return (false, null, 0, "DirError");
            }

            string safeBase = Path.GetFileNameWithoutExtension(baseName.Trim());
            string finalPath = Path.Combine(folder, safeBase + extension);

            string tmpCpd = Path.Combine(Path.GetTempPath(), $"calcpad_{DateTime.Now.Ticks}.cpd");
            try
            {
                File.WriteAllText(tmpCpd, sheet.OriginalCode, new UTF8Encoding(false));
                var parser = new Parser { Settings = new Settings { Math = new MathSettings { Decimals = 15 } } };
                log?.Invoke($"Calcpad Convert → {extension.ToUpperInvariant().Trim('.')} …");
                bool ok = parser.Convert(tmpCpd, finalPath);
                try { File.Delete(tmpCpd); } catch { }

                if (ok && File.Exists(finalPath))
                {
                    long size = new FileInfo(finalPath).Length;
                    return (true, finalPath, size, "Calcpad Convert");
                }
                return (false, finalPath, 0, "ConvertFailed");
            }
            catch (Exception ex)
            {
                try { File.Delete(tmpCpd); } catch { }
                // Improve error detail (InnerException if coming from reflection)
                string msg = (ex is TargetInvocationException tie && tie.InnerException != null)
                    ? $"{tie.InnerException.GetType().Name}: {tie.InnerException.Message}"
                    : ex.Message;

                log?.Invoke($"Native Convert failure: {msg}");
                return (false, finalPath, 0, "Exception");
            }
        }

        // Word Interop in STA with OMaths.BuildUp
        private static (bool success, long size) BuildDocxWithOMathSTA(CalcpadSheet sheet, string docxPath, Action<string> log)
        {
            bool success = false;
            long size = 0;
            Exception threadEx = null;

            Thread t = new Thread(() =>
            {
                object wordApp = null, docs = null, doc = null;
                try
                {
                    var wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        log?.Invoke("Microsoft Word is not installed (ProgID 'Word.Application' missing).");
                        return;
                    }

                    wordApp = Activator.CreateInstance(wordType);
                    wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { false });

                    docs = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    doc = docs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, docs, null);

                    // Title
                    AppendParagraph(doc, $"Calcpad Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}", bold: true, fontName: "Times New Roman", fontSizePt: 14f);

                    // Result equations
                    var eqs = sheet.GetResultEquations();
                    var lhsSet = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (eqs != null)
                        foreach (var eq in eqs)
                        {
                            var lhs0 = ExtractLhs(eq);
                            if (!string.IsNullOrEmpty(lhs0)) lhsSet.Add(lhs0);
                        }

                    var addedLhs = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // Source code traversal
                    var lines = sheet.OriginalCode.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                    foreach (var raw in lines)
                    {
                        var line = (raw ?? string.Empty).TrimEnd();
                        if (string.IsNullOrWhiteSpace(line)) { AppendEmptyParagraph(doc); continue; }
                        if (IsCommentLine(line)) { AppendParagraph(doc, line, italic: true, fontName: "Times New Roman", fontSizePt: 11f); continue; }

                        int eqPos = line.IndexOf('=');
                        if (eqPos > 0)
                        {
                            string lhs = line.Substring(0, eqPos).Trim();
                            if (lhsSet.Contains(lhs))
                            {
                                string rhs = RemoveInlineComments(line.Substring(eqPos + 1));
                                string linear = lhs + " = " + rhs;
                                AppendEquation(doc, linear);
                                addedLhs.Add(lhs);
                                continue;
                            }
                        }
                        AppendParagraph(doc, line, fontName: "Times New Roman", fontSizePt: 12f);
                    }

                    // Results block
                    if (eqs != null && eqs.Count > 0)
                    {
                        AppendEmptyParagraph(doc);
                        AppendParagraph(doc, "Results", bold: true, fontName: "Times New Roman", fontSizePt: 12f);

                        foreach (var eq in eqs)
                        {
                            var lhs = ExtractLhs(eq);
                            if (!string.IsNullOrEmpty(lhs) && addedLhs.Contains(lhs)) continue;
                            AppendEquation(doc, eq);
                        }
                    }

                    // Save
                    doc.GetType().InvokeMember("SaveAs2", BindingFlags.InvokeMethod, null, doc, new object[] { docxPath, 12 });
                    doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] { false });
                    wordType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, wordApp, null);

                    if (File.Exists(docxPath))
                    {
                        size = new FileInfo(docxPath).Length;
                        success = true;
                    }
                }
                catch (Exception ex)
                {
                    var msg = (ex is TargetInvocationException tie && tie.InnerException != null)
                        ? $"{tie.InnerException.GetType().Name}: {tie.InnerException.Message}"
                        : ex.Message;
                    threadEx = new Exception(msg, ex);

                    try { if (doc != null) doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] { false }); } catch { }
                    try { if (wordApp != null) wordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, wordApp, null); } catch { }
                }
                finally
                {
                    ReleaseCom(doc); ReleaseCom(docs); ReleaseCom(wordApp);
                }
            });

            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            if (threadEx != null)
                log?.Invoke($"Word Interop error: {threadEx.Message}");

            return (success, size);
        }

        // ===== Interop / text helpers =====

        private static object GetEndRange(object doc)
        {
            var content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
            int end = (int)content.GetType().InvokeMember("End", BindingFlags.GetProperty, null, content, null);
            var endRange = doc.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, doc, new object[] { end, end });
            return endRange;
        }

        // Inserts a paragraph of text at the end (via Paragraphs.Add)
        private static void AppendParagraph(object doc, string text, bool bold = false, bool italic = false, string fontName = "Times New Roman", float fontSizePt = 12f)
        {
            var paragraphs = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
            var par = paragraphs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, paragraphs, null);
            var range = par.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, par, null);

            range.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, range, new object[] { text });

            var font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
            font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { fontName });
            font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { fontSizePt });
            font.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, font, new object[] { bold ? 1 : 0 });
            font.GetType().InvokeMember("Italic", BindingFlags.SetProperty, null, font, new object[] { italic ? 1 : 0 });

            // Ensure paragraph break
            range.GetType().InvokeMember("InsertParagraphAfter", BindingFlags.InvokeMethod, null, range, null);
        }

        private static void AppendEmptyParagraph(object doc)
        {
            var paragraphs = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
            paragraphs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, paragraphs, null);
        }

        // Converts an equation into OMath without including the paragraph CR in the range (avoid 'out of range')
        private static void AppendEquation(object doc, string linearEq)
        {
            string eq = SanitizeForOMath(linearEq);
            if (string.IsNullOrWhiteSpace(eq) || eq.IndexOf('=') < 0)
            {
                AppendParagraph(doc, linearEq, fontName: "Times New Roman", fontSizePt: 12f);
                return;
            }

            var paragraphs = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
            var par = paragraphs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, paragraphs, null);
            var parRange = par.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, par, null);

            // Write equation text
            parRange.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, parRange, new object[] { eq });

            // Compute a range excluding trailing paragraph mark
            int start = (int)parRange.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, parRange, null);
            int end = (int)parRange.GetType().InvokeMember("End", BindingFlags.GetProperty, null, parRange, null);
            if (end > start) end -= 1; // exclude CR

            var eqRange = doc.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, doc, new object[] { start, end });

            try
            {
                // Math font
                var font = eqRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, eqRange, null);
                font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { "Cambria Math" });
                font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { 12f });

                // Convert to OMath
                var omaths = eqRange.GetType().InvokeMember("OMaths", BindingFlags.GetProperty, null, eqRange, null);
                omaths.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, omaths, new object[] { eqRange });
                omaths.GetType().InvokeMember("BuildUp", BindingFlags.InvokeMethod, null, omaths, null);
            }
            catch
            {
                // Fallback: leave as plain text if OMath conversion fails
                var f = eqRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, eqRange, null);
                f.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, f, new object[] { "Times New Roman" });
                f.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, f, new object[] { 12f });
            }
            finally
            {
                // Ensure paragraph break
                eqRange.GetType().InvokeMember("InsertParagraphAfter", BindingFlags.InvokeMethod, null, eqRange, null);
            }
        }

        private static string SanitizeForOMath(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var t = s;

            // Invisible spaces to normal space
            t = t.Replace('\u2009', ' ').Replace('\u200A', ' ').Replace('\u202F', ' ')
                 .Replace('\u00A0', ' ').Replace('\u2002', ' ').Replace('\u2003', ' ')
                 .Replace('\u2005', ' ').Replace('\u2006', ' ');

            // Common operators
            t = t.Replace('·', '*').Replace('×', '*');

            // Simple powers / roots
            t = t.Replace("²", "^2").Replace("³", "^3").Replace("√", "sqrt ");

            // Final trim
            t = t.Trim();
            return t;
        }

        private static void SetRangeFont(object range, string name, float sizePt)
        {
            var font = range.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, range, null);
            font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { name });
            font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { sizePt });
        }

        private static void ReleaseCom(object x)
        {
            try { if (x != null && Marshal.IsComObject(x)) Marshal.ReleaseComObject(x); } catch { }
        }

        private static bool IsCommentLine(string line)
        {
            var t = (line ?? string.Empty).TrimStart();
            return t.StartsWith("#") || t.StartsWith("'") || t.StartsWith("’") || t.StartsWith("‘");
        }

        private static string ExtractLhs(string eq)
        {
            int i = eq.IndexOf('=');
            if (i <= 0) return string.Empty;
            return eq.Substring(0, i).Trim();
        }

        private static string RemoveInlineComments(string rhs)
        {
            if (string.IsNullOrEmpty(rhs)) return rhs;
            var s = rhs.Replace('\u2009', ' ').Replace('\u200A', ' ').Replace('\u202F', ' ')
                       .Replace('\u00A0', ' ').Replace('\u2002', ' ').Replace('\u2003', ' ')
                       .Replace('\u2005', ' ').Replace('\u2006', ' ');
            s = s.Trim();
            int p = s.IndexOf("#", StringComparison.Ordinal);
            int q = s.IndexOf('\'');
            int cut = s.Length;
            if (p >= 0) cut = Math.Min(cut, p);
            if (q >= 0) cut = Math.Min(cut, q);
            return s.Substring(0, cut).Trim();
        }

        private static string ResolveCalcpadCliPath(string cliPath, Action<string> log)
        {
            // 1) Provided path
            if (!string.IsNullOrWhiteSpace(cliPath) && File.Exists(cliPath)) return cliPath;

            // 2) Environment variable
            var env = Environment.GetEnvironmentVariable("CALCPAD_CLI");
            if (!string.IsNullOrWhiteSpace(env) && File.Exists(env)) return env;

            // 3) Next to plugin
            var here = Path.GetDirectoryName(typeof(CalcpadExporter).Assembly.Location);
            if (!string.IsNullOrEmpty(here))
            {
                var local = Path.Combine(here, "Calcpad.Cli.exe");
                if (File.Exists(local)) return local;
            }

            // 4) Installed program (heuristic)
            var pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var guess = Path.Combine(pf, "Calcpad", "Calcpad.Cli.exe");
            if (File.Exists(guess)) return guess;

            return null;
        }

        private static (bool ok, string error) RunProcess(string exe, string args, Action<string> log, int timeoutMs)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                using (var p = new Process { StartInfo = psi })
                {
                    var sbOut = new StringBuilder();
                    var sbErr = new StringBuilder();
                    p.OutputDataReceived += (s, e) => { if (e.Data != null) sbOut.AppendLine(e.Data); };
                    p.ErrorDataReceived += (s, e) => { if (e.Data != null) sbErr.AppendLine(e.Data); };

                    if (!p.Start()) return (false, "Could not start CLI process.");
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();

                    if (!p.WaitForExit(timeoutMs))
                    {
                        try { p.Kill(); } catch { }
                        return (false, "CLI execution timeout.");
                    }

                    string allOut = sbOut.ToString().Trim();
                    string allErr = sbErr.ToString().Trim();
                    if (!string.IsNullOrEmpty(allOut)) log?.Invoke(allOut);
                    if (!string.IsNullOrEmpty(allErr)) log?.Invoke(allErr);

                    return (p.ExitCode == 0, p.ExitCode == 0 ? null : $"ExitCode={p.ExitCode} {allErr}");
                }
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private static void TryDelete(string path)
        {
            try { if (!string.IsNullOrEmpty(path) && File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}