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

            // Enrich unitChars/functions from the embedded calcpad.xml (AutoComplete/KeyWord format)
            TryAugmentUnitCharsFromEmbeddedCalcpadXml();

            // Build regex with current unit character set
            RebuildRegexFromUnitChars();
        }

        public void ParseVariables(string content, out List<string> names, out List<double> values, out List<string> units)
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

        // A leading numeric literal followed only by a compound unit (letters, °/µ/Ω/·,
        // a single "/" for rate units like kN/m, and "^n" exponents like m^3 or kg/m^3)
        // is a plain value assignment, not an equation - even though "/" and "^" are also
        // arithmetic operators. Only flag it as an equation if something is left over that
        // a unit token can't explain: +, -, *, parentheses, or a second bare number.
        private static readonly Regex RxLeadingNumber = new Regex(
            @"^[-+]?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?", RegexOptions.Compiled);
        private static readonly Regex RxUnitExponent = new Regex(
            @"\^-?\d+", RegexOptions.Compiled);

        private bool LooksLikeEquation(string rhs)
        {
            if (string.IsNullOrWhiteSpace(rhs)) return false;
            rhs = rhs.TrimStart();
            if (rhs.StartsWith("?", StringComparison.Ordinal)) return false;

            var lead = RxLeadingNumber.Match(rhs);
            string tail = lead.Success ? rhs.Substring(lead.Length) : rhs;
            tail = RxUnitExponent.Replace(tail, "");

            if (Regex.IsMatch(tail, @"[+\-*()]")) return true;
            if (Regex.IsMatch(tail, @"\d")) return true; // leftover digit not part of a unit exponent

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
        }

        // Build character class for units from _unitChars and recompile regex
        private void RebuildRegexFromUnitChars()
        {
            var unitClass = BuildUnitCharClass(_unitChars);
            _unitCharClass = unitClass;

            // Identifier class for names, matching Calcpad.Core's own Validator.VarChars set:
            // straight apostrophe ' is NEVER a valid name char in real Calcpad - it is always
            // a comment/quote delimiter - only the real prime marks are (single/double/triple/quadruple).
            const string nameClass = "A-Za-z0-9_,\\.′″‴⁗℧∡ϑϕøØ";

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

    }
}