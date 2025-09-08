using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.Net;
using System.Globalization;
using Grasshopper.Kernel;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    public class GH_Calcpad_Load_cpdz : GH_Component
    {
        private FileSystemWatcher _watcher;

        public GH_Calcpad_Load_cpdz()
          : base("Load CPDz", "LoadCPDz",
                 "Reads a .cpdz (ZIP/Base64). If a textual .cpd is present, parses it; otherwise tries to decode nested/compressed payloads. Returns only explicit variables (?{...}).",
                 "Calcpad", "2. File Loading")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter("FilePath", "P", "Complete path to .cpdz file", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("Variables", "N", "Names of explicit variables found", GH_ParamAccess.list);
            p.AddNumberParameter("Values", "V", "Corresponding values", GH_ParamAccess.list);
            p.AddTextParameter("Units", "U", "Corresponding units", GH_ParamAccess.list);
            p.AddGenericParameter("SheetObj", "S", "CalcpadSheet instance", GH_ParamAccess.item);
            p.AddTextParameter("Status", "St", "TextOK | Binary | NoSource | HtmlParsed", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string path = null;
            if (!DA.GetData(0, ref path)) return;

            SetupFileWatcher(path);

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $".cpdz file does not exist:\n{path}");
                return;
            }

            string sourceText;
            List<string> zipEntries;
            string status;
            try
            {
                sourceText = ExtractBestTextFromCpdz(path, out zipEntries, out status);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error unpacking .cpdz: {ex.Message}");
                return;
            }

            if (status == "Binary")
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    "This .cpdz contains a compiled/binary code.cpd (no textual source). Cannot extract variables. Export the package including the textual .cpd or provide the .cpd file."
                );
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetDataList(2, new List<string>());
                DA.SetData(3, null);
                DA.SetData(4, status);
                return;
            }

            if (string.IsNullOrWhiteSpace(sourceText))
            {
                if (status != "Binary")
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty text extracted from .cpdz.");
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetDataList(2, new List<string>());
                DA.SetData(3, null);
                DA.SetData(4, string.IsNullOrEmpty(status) ? "NoSource" : status);
                return;
            }

            // Normalization before parsing
            var normalized = NormalizeForParser(sourceText);

            // 1) Explicit only (yellow boxes)
            List<string> names, units;
            List<double> values;
            CalcpadSyntax.Instance.ParseVariables(normalized, captureExplicit: true, out names, out values, out units);

            // 2) Brief message if no explicit variables found
            if (names.Count == 0)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "No explicit ?{…} inputs found in this package.");

            var sheetObj = new CalcpadSheet(names, values, units);
            try { sheetObj.SetFullCode(normalized); }
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
            DA.SetData(4, status);
        }

        // Extract best possible text and classify Status.
        // Status: "TextOK" (textual cpd) | "Binary" (compiled cpd) | "HtmlParsed" (text from html) | "NoSource" (no useful source)
        private static string ExtractBestTextFromCpdz(string path, out List<string> entriesOut, out string status)
        {
            status = "NoSource";
            byte[] fileBytes = File.ReadAllBytes(path);

            // Some .cpdz files come in Base64
            byte[] zipBytes;
            try
            {
                string fileContent = Encoding.UTF8.GetString(fileBytes);
                zipBytes = Convert.FromBase64String(fileContent);
            }
            catch (FormatException)
            {
                zipBytes = fileBytes;
            }

            using (var ms = new MemoryStream(zipBytes))
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                entriesOut = archive.Entries.Select(e => e.FullName).ToList();

                // 1) Prefer .cpd
                var cpdEntry = archive.Entries
                    .FirstOrDefault(e => e.Name.EndsWith(".cpd", StringComparison.OrdinalIgnoreCase));
                if (cpdEntry != null)
                {
                    string text = ReadZipEntrySmart(cpdEntry, out bool cpdBinary);
                    if (cpdBinary)
                    {
                        status = "Binary";
                        return string.Empty;
                    }
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        status = "TextOK";
                        return text;
                    }
                }

                // 2) HTML/TXT
                var textCandidates = archive.Entries
                    .Where(e =>
                        e.Name.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.EndsWith(".md", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.EndsWith(".htm", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(e => e.FullName)
                    .ToList();

                bool anyHtml = false;
                var sb = new StringBuilder();
                foreach (var entry in textCandidates)
                {
                    var raw = ReadZipEntryTextWithEncoding(entry);
                    bool isHtml = entry.Name.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                                  entry.Name.EndsWith(".htm", StringComparison.OrdinalIgnoreCase);
                    anyHtml |= isHtml;
                    var plain = isHtml ? HtmlToPlain(raw) : raw;
                    if (!string.IsNullOrWhiteSpace(plain))
                    {
                        sb.AppendLine(plain);
                        sb.AppendLine();
                    }
                }

                string combined = sb.ToString();
                if (!string.IsNullOrWhiteSpace(combined))
                {
                    status = anyHtml ? "HtmlParsed" : "TextOK";
                    return combined;
                }

                status = "NoSource";
                return string.Empty;
            }
        }

        // Detect text/gzip/deflate/nested zip. If not parseable, mark binary=true.
        private static string ReadZipEntrySmart(ZipArchiveEntry entry, out bool binary)
        {
            binary = false;

            using (var es = entry.Open())
            using (var ms = new MemoryStream())
            {
                es.CopyTo(ms);
                var data = ms.ToArray();

                // Nested zip (PK..)
                if (data.Length > 4 && data[0] == 0x50 && data[1] == 0x4B)
                {
                    using (var inner = new MemoryStream(data))
                    using (var innerZip = new ZipArchive(inner, ZipArchiveMode.Read))
                    {
                        var innerCpd = innerZip.Entries.FirstOrDefault(e => e.Name.EndsWith(".cpd", StringComparison.OrdinalIgnoreCase));
                        if (innerCpd != null)
                        {
                            string txt = ReadZipEntryTextWithEncoding(innerCpd);
                            if (!string.IsNullOrWhiteSpace(txt)) return txt;
                            binary = true; // also came compiled inside
                            return string.Empty;
                        }

                        var html = innerZip.Entries.FirstOrDefault(e =>
                            e.Name.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                            e.Name.EndsWith(".htm", StringComparison.OrdinalIgnoreCase));
                        if (html != null)
                            return HtmlToPlain(ReadZipEntryTextWithEncoding(html));
                    }
                }

                // GZip (1F 8B)
                if (data.Length > 2 && data[0] == 0x1F && data[1] == 0x8B)
                {
                    using (var ims = new MemoryStream(data))
                    using (var gz = new GZipStream(ims, CompressionMode.Decompress))
                    using (var outMs = new MemoryStream())
                    {
                        gz.CopyTo(outMs);
                        return DecodeTextFromBytes(outMs.ToArray());
                    }
                }

                // zlib/deflate (78 01 / 78 9C / 78 DA)
                if (data.Length > 2 && data[0] == 0x78 && (data[1] == 0x01 || data[1] == 0x9C || data[1] == 0xDA))
                {
                    using (var ims = new MemoryStream(data))
                    using (var def = new DeflateStream(ims, CompressionMode.Decompress))
                    using (var outMs = new MemoryStream())
                    {
                        def.CopyTo(outMs);
                        return DecodeTextFromBytes(outMs.ToArray());
                    }
                }

                // Looks like text?
                if (IsLikelyText(data))
                {
                    return DecodeTextFromBytes(data);
                }

                // Binary/compiled
                binary = true;
                return string.Empty;
            }
        }

        private static string DecodeTextFromBytes(byte[] bytes)
        {
            var enc = DetectEncoding(bytes) ?? Encoding.UTF8;
            string text = enc.GetString(bytes);
            if (!string.IsNullOrEmpty(text) && text[0] == '\uFEFF') text = text.Substring(1);
            return text.Replace("\r\n", "\n");
        }

        private static bool IsLikelyText(byte[] data)
        {
            if (data == null || data.Length == 0) return false;
            int control = 0, printable = 0, probe = Math.Min(data.Length, 4096);
            for (int i = 0; i < probe; i++)
            {
                byte b = data[i];
                if (b == 0x09 || b == 0x0A || b == 0x0D) { printable++; continue; }
                if (b >= 0x20 && b <= 0x7E) { printable++; continue; }
                if (b >= 0xC2 && b <= 0xF4) { printable++; continue; } // UTF-8 multibyte start (heuristic)
                control++;
            }
            return printable >= control * 4;
        }

        private static string ReadZipEntryTextWithEncoding(ZipArchiveEntry entry)
        {
            using (var es = entry.Open())
            using (var ms = new MemoryStream())
            {
                es.CopyTo(ms);
                return DecodeTextFromBytes(ms.ToArray());
            }
        }

        private static Encoding DetectEncoding(byte[] data)
        {
            if (data == null || data.Length < 2) return null;

            if (data.Length >= 4)
            {
                if (data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00 && data[3] == 0x00) return Encoding.UTF32; // LE
                if (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF) return new UTF32Encoding(true, true); // BE
            }
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) return Encoding.UTF8;
            if (data[0] == 0xFF && data[1] == 0xFE) return Encoding.Unicode;           // UTF-16 LE
            if (data[0] == 0xFE && data[1] == 0xFF) return Encoding.BigEndianUnicode;  // UTF-16 BE

            int zerosEven = 0, zerosOdd = 0, pairs = data.Length / 2;
            for (int i = 0; i + 1 < data.Length; i += 2)
            {
                if (data[i] == 0x00) zerosEven++;
                if (data[i + 1] == 0x00) zerosOdd++;
            }
            if (pairs > 0)
            {
                double rEven = zerosEven / (double)pairs;
                double rOdd = zerosOdd / (double)pairs;
                if (rOdd > 0.3 && rOdd > rEven) return Encoding.Unicode;
                if (rEven > 0.3 && rEven > rOdd) return Encoding.BigEndianUnicode;
            }

            return Encoding.UTF8;
        }

        private static string NormalizeForParser(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;

            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch == '\r') continue;
                if (ch == '\n') { sb.Append('\n'); continue; }

                var cat = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (cat == UnicodeCategory.Format) continue; // U+200B, BOM, etc.

                if (char.IsWhiteSpace(ch)) sb.Append(' ');
                else sb.Append(ch);
            }

            return Regex.Replace(sb.ToString(), @"[ ]{2,}", " ");
        }

        private static string HtmlToPlain(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;
            string s = Regex.Replace(html, "(?is)<script.*?</script>", "");
            s = Regex.Replace(s, "(?is)<style.*?</style>", "");
            s = Regex.Replace(s, "(?i)</p>|</div>|<br\\s*/?>", "\n");
            s = Regex.Replace(s, "<[^>]+>", " ");
            s = WebUtility.HtmlDecode(s);
            s = Regex.Replace(s, "\\s+", " ").Trim();
            return s;
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
                _watcher.Changed += (s, e) => ExpireSolution(true);
            }
            catch { }
        }

        public override Guid ComponentGuid => new Guid("F2E1D8C7-B6A5-4F3E-8D7C-1A2B3C4D5E6F");
        protected override System.Drawing.Bitmap Icon => Resources.Icon_Form;
        public override GH_Exposure Exposure => GH_Exposure.primary;

        public override void RemovedFromDocument(Grasshopper.Kernel.GH_Document document)
        {
            try { _watcher?.Dispose(); _watcher = null; } catch { }
            base.RemovedFromDocument(document);
        }
    }
}