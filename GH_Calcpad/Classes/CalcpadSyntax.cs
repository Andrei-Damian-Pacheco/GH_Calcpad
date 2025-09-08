using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace GH_Calcpad.Classes
{
    /// <summary>
    /// Calcpad syntax loader/validator with fallback to internal sets.
    /// Provides robust variable extraction (explicit and literal) avoiding equations.
    /// </summary>
    public sealed class CalcpadSyntax
    {
        public static CalcpadSyntax Instance { get; } = new CalcpadSyntax();

        private readonly HashSet<string> _functions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _keywords  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Dynamic set of valid characters in "unit"
        private readonly HashSet<char> _unitChars = new HashSet<char>(new[]
        {
            // base
            'A','B','C','D','E','F','G','H','I','J','K','L','M',
            'N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
            'a','b','c','d','e','f','g','h','i','j','k','l','m',
            'n','o','p','q','r','s','t','u','v','w','x','y','z',
            '0','1','2','3','4','5','6','7','8','9',
            '%','/','^','-','.','*','(',')','_',' ',
            // typical symbols that appear in calcpad.xml
            'µ', // MICRO SIGN U+00B5
            'μ', // GREEK SMALL LETTER MU U+03BC
            '°', // DEGREE SIGN U+00B0
            'Ω', // OHM SIGN/GREEK OMEGA U+03A9
            '℧', // MHO SIGN U+2127
            'Δ', // GREEK CAPITAL LETTER DELTA U+0394
            '·'  // MIDDLE DOT U+00B7 (kN·m)
        });

        // Operators/traits that classify as equation
        private static readonly char[] _opChars = new[] { '+', '-', '*', '/', '^', '(', ')', '=' };

        // Dynamically compilable regex (rebuilt after loading XML)
        private Regex RxExplicit1;
        private Regex RxExplicit2Inline;
        private Regex RxLiteralAssign;

        // NEW: character class for units (safe list for inside [])
        private string _unitCharClass = string.Empty;
        public string UnitCharClass => _unitCharClass;

        private CalcpadSyntax()
        {
            // Quick fallback
            SeedFallback();

            // Try to enrich unitChars from calcpad.xml (AutoComplete/KeyWord format)
            TryAugmentUnitCharsFromCalcpadXmlNearAssembly();
            TryAugmentUnitCharsFromEmbeddedCalcpadXml();

            // Build regex with current unit character set
            RebuildRegexFromUnitChars();

            // Load functions from Notepad++ XML files (if they exist)
            TryLoadFromNotepadPlusPlusSyntax();
        }

        public void ParseVariables(string content, bool captureExplicit, out List<string> names, out List<double> values, out List<string> units)
        {
            names = new List<string>();
            values = new List<double>();
            units  = new List<string>();
            if (string.IsNullOrWhiteSpace(content)) return;

            var map = new Dictionary<string, (double val, string unit)>(StringComparer.OrdinalIgnoreCase);
            var order = new List<string>();

            using (var reader = new StringReader(content))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // 0) Capture ALL valueunit';'name cases in the line (multiple possible)
                    //    Avoid confusion with "…';'name = …" via negative lookahead (not followed by '=').
                    var inlineMatches = RxExplicit2Inline.Matches(line);
                    foreach (Match m in inlineMatches)
                    {
                        string name = m.Groups["name"].Value.Trim();
                        string unit = NormalizeUnit(m.Groups["unit"].Value);
                        if (TryParseDouble(m.Groups["val"].Value, out var dv))
                            Upsert(map, order, name, dv, unit);
                    }

                    // 1) Replace "';'" with ';' and remove real comments (' and #)
                    var noComments = StripCommentsPreservingSeparator(line);

                    // 2) Split by ';' to support multiple assignments per line
                    foreach (var segment in SplitSegments(noComments))
                    {
                        // var = ?{val}unit (explicit)
                        var m1 = RxExplicit1.Match(segment);
                        if (m1.Success)
                        {
                            string name = m1.Groups["name"].Value.Trim();
                            string unit = NormalizeUnit(m1.Groups["unit"].Value);
                            if (TryParseDouble(m1.Groups["val"].Value, out var dv))
                                Upsert(map, order, name, dv, unit);
                            continue;
                        }

                        if (captureExplicit) continue;

                        // Literals: name = 123 [unit] (multiple can exist per segment)
                        var m3s = RxLiteralAssign.Matches(segment);
                        foreach (Match m3 in m3s)
                        {
                            var rhs = ExtractRhs(m3.Value);
                            if (LooksLikeEquation(rhs)) continue;

                            string name = m3.Groups["name"].Value.Trim();
                            string unit = NormalizeUnit(m3.Groups["unit"].Value);
                            if (TryParseDouble(m3.Groups["num"].Value, out var dv))
                                Upsert(map, order, name, dv, unit);
                        }
                    }
                }
            }

            foreach (var key in order)
            {
                var tup = map[key];
                names.Add(key);
                values.Add(tup.val);
                units.Add(tup.unit);
            }
        }

        private static void Upsert(Dictionary<string,(double val,string unit)> map, List<string> order, string name, double val, string unit)
        {
            if (!map.ContainsKey(name)) order.Add(name);
            map[name] = (val, unit);
        }

        private static IEnumerable<string> SplitSegments(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) yield break;
            var parts = s.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var seg = p.Trim();
                if (seg.Length > 0) yield return seg;
            }
        }

        // Collapse any Unicode whitespace to 1 space and trim
        private static string NormalizeUnit(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit)) return string.Empty;
            unit = Regex.Replace(unit, @"\s+", " ");
            return unit.Trim();
        }

        // Preserve "';'" as statement separator and only remove real comments with ' or #
        private static string StripCommentsPreservingSeparator(string line)
        {
            if (string.IsNullOrEmpty(line)) return line;

            // Remove # ... end of line
            int iHash = line.IndexOf('#');
            if (iHash >= 0) line = line.Substring(0, iHash);

            var sb = new StringBuilder(line.Length);
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\'')
                {
                    // If it's the "';'" token, convert to ';' and continue
                    if (i + 2 < line.Length && line[i + 1] == ';' && line[i + 2] == '\'')
                    {
                        sb.Append(';');
                        i += 2; // skip ;'
                        continue;
                    }
                    // Otherwise, it's a Calcpad comment: cut here
                    break;
                }
                sb.Append(c);
            }
            return sb.ToString().TrimEnd();
        }

        private static string ExtractRhs(string line)
        {
            int idx = line.IndexOf('=');
            if (idx < 0) return line;
            return line.Substring(idx + 1).Trim();
        }

        private bool LooksLikeEquation(string rhs)
        {
            if (string.IsNullOrWhiteSpace(rhs)) return false;
            if (rhs.TrimStart().StartsWith("?", StringComparison.Ordinal)) return false;
            if (rhs.IndexOfAny(_opChars) >= 0) return true;

            foreach (var fn in _functions)
            {
                if (rhs.IndexOf(fn, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (Regex.IsMatch(rhs, $@"\b{Regex.Escape(fn)}\s*\(", RegexOptions.IgnoreCase))
                        return true;
                }
            }
            return false;
        }

        private static bool TryParseDouble(string s, out double d)
        {
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        private void SeedFallback()
        {
            var minFunctions = new[]
            {
                "sqrt","sin","cos","tan","asin","acos","atan",
                "exp","ln","log","abs","min","max","round","floor","ceil",
                "sinh","cosh","tanh","pow"
            };
            foreach (var f in minFunctions) _functions.Add(f);

            var minKeywords = new[]
            {
                "for","to","step","if","else","endif","endfor","while","endwhile",
                "table","print","plot","const"
            };
            foreach (var k in minKeywords) _keywords.Add(k);
        }

        // Build character class for units from _unitChars and recompile regex
        private void RebuildRegexFromUnitChars()
        {
            var unitClass = BuildUnitCharClass(_unitChars);
            _unitCharClass = unitClass;

            // identifier class for names (same set as name pattern)
            const string nameClass = "A-Za-z0-9_'′,\\.";

            // var = ?{value}unit (anchored to segment)
            RxExplicit1 = new Regex(
                @"^(?<lead>\s*)(?<name>[A-Za-z_][" + nameClass + @"]*)\s*=\s*\?\s*\{\s*(?<val>[-+]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)\s*\}(?<unit>[" + unitClass + @"]*)$",
                RegexOptions.Compiled | RegexOptions.Multiline);

            // valueunit';'name inline (multiple), WITHOUT allowing identifier cut
            RxExplicit2Inline = new Regex(
                @"(?<!\S)(?<val>[-+]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)\s*(?<unit>[" + unitClass + @"]*?)\s*';'\s*(?<name>[A-Za-z_][" + nameClass + @"]*)(?![" + nameClass + @"])(?!\s*=)",
                RegexOptions.Compiled);

            // name = 123 [unit] (not anchored for Matches)
            RxLiteralAssign = new Regex(
                @"(?<!\S)(?<name>[A-Za-z_][" + nameClass + @"]*)\s*=\s*(?<num>[-+]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)(?<unit>\s*[" + unitClass + @"]*)",
                RegexOptions.Compiled);
        }

        // Convert char set to escaped regex class (inside [])
        private static string BuildUnitCharClass(HashSet<char> chars)
        {
            var sb = new StringBuilder();
            foreach (var ch in chars)
            {
                // Escape class specials: \ - ] ^
                if (ch == '\\' || ch == '-' || ch == ']' || ch == '^')
                    sb.Append('\\').Append(ch);
                else
                    sb.Append(ch);
            }
            return sb.ToString();
        }

        // Read calcpad.xml in assembly folder and add unique token characters to _unitChars
        private void TryAugmentUnitCharsFromCalcpadXmlNearAssembly()
        {
            try
            {
                var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
                var candidates = new[]
                {
                    Path.Combine(asmDir, "calcpad.xml"),
                    Path.Combine(asmDir, "Syntaxis", "calcpad.xml"),
                    Path.Combine(asmDir, "Syntax", "calcpad.xml"),
                    Path.Combine(asmDir, "Documents", "calcpad.xml"),
                };
                foreach (var path in candidates)
                {
                    if (File.Exists(path))
                    {
                        AugmentUnitCharsFromAutoCompleteXml(path);
                        break;
                    }
                }
            }
            catch { /* fallback */ }
        }

        // Read embedded calcpad.xml as resource and add token characters to _unitChars
        private void TryAugmentUnitCharsFromEmbeddedCalcpadXml()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                foreach (var res in asm.GetManifestResourceNames())
                {
                    if (res.EndsWith("calcpad.xml", StringComparison.OrdinalIgnoreCase))
                    {
                        using (var s = asm.GetManifestResourceStream(res))
                        {
                            if (s == null) continue;
                            var doc = new XmlDocument();
                            doc.Load(s);
                            AugmentUnitCharsFromAutoCompleteXml(doc);
                        }
                    }
                }
            }
            catch { /* silent */ }
        }

        // Extract <AutoComplete><KeyWord name="..."/> and add non-alphanumeric characters useful for units
        private void AugmentUnitCharsFromAutoCompleteXml(string xmlPath)
        {
            try
            {
                var doc = new XmlDocument();
                doc.Load(xmlPath);
                AugmentUnitCharsFromAutoCompleteXml(doc);
            }
            catch { /* ignore parse errors */ }
        }

        private void AugmentUnitCharsFromAutoCompleteXml(XmlDocument doc)
        {
            try
            {
                var nodes = doc.SelectNodes("//NotepadPlus/AutoComplete/KeyWord[@name]");
                if (nodes == null) return;

                foreach (XmlNode n in nodes)
                {
                    var tok = n.Attributes?["name"]?.Value ?? "";
                    if (string.IsNullOrWhiteSpace(tok)) continue;

                    tok = tok.Trim(); // there are tokens with trailing space in XML (e.g. "kip ")
                    foreach (var ch in tok)
                    {
                        // Add only characters potentially present in units.
                        if (char.IsLetterOrDigit(ch)) { /* already there */ }
                        else if (!char.IsControl(ch))
                        {
                            _unitChars.Add(ch);
                        }
                    }

                    // Also leverage for known functions (heuristic improvement)
                    if (Regex.IsMatch(tok, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                        _functions.Add(tok);
                }
            }
            catch { /* ignore parse errors */ }
        }

        private void TryLoadFromNotepadPlusPlusSyntax()
        {
            try
            {
                var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
                var candidates = new[]
                {
                    Path.Combine(asmDir, "Calcpad-syntax-for-Notepad++.xml"),
                    Path.Combine(asmDir, "Calcpad-syntax-for-Notepad++dark.xml"),
                    Path.Combine(asmDir, "Syntax", "Calcpad-syntax-for-Notepad++.xml"),
                    Path.Combine(asmDir, "Syntax", "Calcpad-syntax-for-Notepad++dark.xml"),
                    Path.Combine(asmDir, "Syntaxis", "Calcpad-syntax-for-Notepad++.xml"),
                    Path.Combine(asmDir, "Syntaxis", "Calcpad-syntax-for-Notepad++dark.xml"),
                    Path.Combine(asmDir, "Documents", "Calcpad-syntax-for-Notepad++.xml"),
                    Path.Combine(asmDir, "Documents", "Calcpad-syntax-for-Notepad++dark.xml"),
                };
                foreach (var path in candidates)
                {
                    if (File.Exists(path)) { ExtractNotepadPlusPlusKeywords(path); }
                }
            }
            catch { /* fallback */ }
        }

        private void ExtractNotepadPlusPlusKeywords(string xmlPath)
        {
            try
            {
                var doc = new XmlDocument();
                doc.Load(xmlPath);
                var nodes = doc.SelectNodes("//Keywords");
                if (nodes != null)
                {
                    foreach (XmlNode n in nodes)
                    {
                        var nameAttr = n.Attributes?["name"]?.Value ?? "";
                        var text = n.InnerText ?? "";
                        foreach (var tok in SplitTokens(text))
                        {
                            if (nameAttr.IndexOf("func", StringComparison.OrdinalIgnoreCase) >= 0)
                                _functions.Add(tok);
                            else
                                _keywords.Add(tok);
                        }
                    }
                }
            }
            catch { /* ignore */ }
        }

        private static IEnumerable<string> SplitTokens(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) yield break;
            foreach (var t in text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var tok = t.Trim();
                if (tok.Length == 0) continue;
                tok = tok.Trim('\'', '"');
                yield return tok;
            }
        }
    }
}