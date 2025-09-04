using PyCalcpad;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace GH_Calcpad.Classes
{
    public class CalcpadSheet
    {
        // Compatibilidad con Load
        public List<string> Variables { get; }
        public List<double> Values { get; }
        public List<string> Units { get; }

        // Estado
        private string _originalCode;
        private Parser _parser;
        private Settings _settings;
        private string _lastHtmlResult;

        private const string UnitTokenClassNoDigits = "A-Za-z°µμΩ℧·/\\-\\^²³";

        public string OriginalCode => _originalCode ?? string.Empty;
        public bool HasCodeAvailable => !string.IsNullOrEmpty(_originalCode);
        public string CodeInfo => string.IsNullOrEmpty(_originalCode) ? "No CPD code" : $"CPD code: {_originalCode.Length} characters";
        public string LastHtmlResult => _lastHtmlResult ?? string.Empty;

        public CalcpadSheet(List<string> variables, List<double> values, List<string> units)
        {
            Variables = variables ?? new List<string>();
            Values = values ?? new List<double>();
            Units = units ?? new List<string>();
            _originalCode = string.Empty;
            _lastHtmlResult = string.Empty;
            InitParser();
        }

        private void InitParser()
        {
            try
            {
                _settings = new Settings();
                _settings.Math.Decimals = 15;

                _parser = new Parser();
                _parser.Settings = _settings;
            }
            catch
            {
                _parser = null;
                _settings = null;
            }
        }

        public void SetFullCode(string code) => _originalCode = code ?? string.Empty;

        public void SetUnit(string name, string unit)
        {
            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code available to modify.");

            try
            {
                string pattern = @"(?m)^(\s*" + Regex.Escape(name) + @"\s*=\s*[0-9\.\-eE]+\s*)([^\r\n]+)?$";
                _originalCode = Regex.Replace(_originalCode, pattern, m => m.Groups[1].Value + unit);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting unit for '{name}': {ex.Message}");
            }
        }

        public void SetVariable(string name, double value)
        {
            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code available.");

            try
            {
                string vStr = value.ToString(CultureInfo.InvariantCulture);

                // Reemplazo por bloque preservando "';'" y unidades
                string pattern = @"(?m)(^|\s*';'\s*)\s*(?<lhs>" + Regex.Escape(name) + @")\s*=\s*(?<rhs>[^\r\n]*?)(?=(\s*';'\s*|$))";

                bool replaced = false;
                _originalCode = Regex.Replace(_originalCode, pattern, m =>
                {
                    if (replaced) return m.Value;
                    string boundary = m.Groups[1].Value;
                    string rhs = m.Groups["rhs"].Value;

                    string unit = ExtractUnitSuffix(rhs);
                    string newRhs = string.IsNullOrEmpty(unit) ? vStr : (vStr + " " + unit);
                    replaced = true;
                    return $"{boundary}{name} = {newRhs}";
                }, RegexOptions.None);

                if (!replaced)
                {
                    string linePattern = @"(?m)^(?<pre>\s*)" + Regex.Escape(name) + @"\s*=\s*(?<rhs>[^\r\n]+)$";
                    _originalCode = Regex.Replace(_originalCode, linePattern, m =>
                    {
                        string pre = m.Groups["pre"].Value;
                        string rhs = m.Groups["rhs"].Value;
                        string unit = ExtractUnitSuffix(rhs);
                        string newRhs = string.IsNullOrEmpty(unit) ? vStr : (vStr + " " + unit);
                        return $"{pre}{name} = {newRhs}";
                    }, RegexOptions.None);
                }

                Debug.WriteLine($"SetVariable('{name}', {value}) - Block-level replace done");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetVariable('{name}', {value}) - ERROR: {ex.Message}");
            }
        }

        public void Calculate()
        {
            if (_parser == null)
                throw new InvalidOperationException("Parser is not available.");
            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code to calculate. Use SetFullCode() first.");

            try
            {
                // Preprocesado opcional de aliases de unidades no soportadas por el parser
                string codeToParse = PreprocessCode(_originalCode);
                _lastHtmlResult = _parser.Parse(codeToParse);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Calculation error: {ex.Message}");
            }
        }

        // ==================== Ecuaciones / Valores / Unidades ====================

        // Independiente de LoadCPD True/False
        public List<string> GetResultEquations()
        {
            var equations = new List<string>();
            try
            {
                var lines = _originalCode.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    string clean = line.Trim();
                    if (string.IsNullOrEmpty(clean)) continue;
                    if (clean.StartsWith("#") || clean.StartsWith("'") || clean.StartsWith("’") || clean.StartsWith("‘")) continue;

                    if (!IsEquationDefinition(clean)) continue;

                    int eq = clean.IndexOf('=');
                    string left = clean.Substring(0, eq).Trim();
                    string right = RemoveInlineComments(clean.Substring(eq + 1).Trim());
                    if (!string.IsNullOrEmpty(right))
                        equations.Add($"{left} = {right}");
                }
            }
            catch { }
            return equations;
        }

        public List<double> GetResultValues()
        {
            var results = new List<double>();
            var vars = GetEquationVariableNamesFromCode(); // orden del código
            if (vars.Count == 0) return results;

            string plain = ToPlainText(_lastHtmlResult);
            var blocks = ExtractResultBlocksByVar(plain, vars);

            foreach (var v in vars)
            {
                if (blocks.TryGetValue(v, out var block))
                    results.Add(ExtractFinalNumericFromBlock(block));
                else
                    results.Add(double.NaN);
            }
            return results;
        }

        public List<string> GetResultUnits()
        {
            var units = new List<string>();
            var vars = GetEquationVariableNamesFromCode();
            if (vars.Count == 0) return units;

            // 1) HTML (reconstruye fracciones dvc/dvl)
            var htmlMap = ExtractResultBlocksByVarHtml(_lastHtmlResult, vars);

            foreach (var v in vars)
            {
                string u = string.Empty;

                if (htmlMap.TryGetValue(v, out var spanHtml))
                    u = ExtractFinalUnitFromEqSpanHtml(spanHtml);

                if (string.IsNullOrEmpty(u))
                {
                    // 2) Fallback: texto plano
                    string plain = ToPlainText(_lastHtmlResult);
                    var textMap = ExtractResultBlocksByVar(plain, vars);
                    if (textMap.TryGetValue(v, out var block))
                        u = ExtractFinalUnitFromBlock(block);
                }

                units.Add(u ?? string.Empty);
            }
            return units;
        }

        // ==================== Reglas de detección ====================

        // Ecuación “real” del código fuente (no asignaciones ni explícitos)
        private bool IsEquationDefinition(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;

            int firstEq = line.IndexOf('=');
            if (firstEq < 0) return false;
            if (line.Contains("';'")) return false;                  // múltiples asignaciones en la línea
            if (line.IndexOf('=', firstEq + 1) >= 0) return false;   // más de un '='

            string left = line.Substring(0, firstEq).Trim();
            string rightRaw = line.Substring(firstEq + 1).Trim();
            string right = RemoveInlineComments(rightRaw);

            // LHS válido
            if (!Regex.IsMatch(left, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$")) return false;

            // Dinámico: clase de caracteres de unidad desde calcpad.xml
            string unitClass = CalcpadSyntax.Instance.UnitCharClass;

            // Excluir RHS puramente numérico (+unidad)
            var rxNumUnit = new Regex(@"^[+-]?\d+(?:\.\d+)?(?:\s*[" + unitClass + @"]+)?\s*$",
                              RegexOptions.CultureInvariant);
            if (rxNumUnit.IsMatch(right)) return false;

            // Excluir explícitos ?{...}[unidad]
            var rxExplicit = new Regex(@"^\?\s*\{\s*[^}]+\s*\}\s*(?:[" + unitClass + @"]+)?\s*$",
                               RegexOptions.CultureInvariant);
            if (rxExplicit.IsMatch(right)) return false;

            // Aceptar SOLO si hay operadores/funciones (verdaderas ecuaciones)
            bool hasOps = right.IndexOfAny(new[] { '+', '-', '*', '/', '^', '(', ')' }) >= 0
               || Regex.IsMatch(right, @"\b(sqrt|sin|cos|tan|log|exp|abs|min|max|pow)\b",
                                RegexOptions.IgnoreCase);

            return hasOps;
        }

        private List<string> GetEquationVariableNamesFromCode()
        {
            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            try
            {
                var lines = _originalCode.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    var clean = line.Trim();
                    if (string.IsNullOrEmpty(clean)) continue;
                    if (clean.StartsWith("#") || clean.StartsWith("'") || clean.StartsWith("’") || clean.StartsWith("‘")) continue;
                    if (!clean.Contains("=") || clean.Contains("';'")) continue;
                    if (!IsEquationDefinition(clean)) continue;

                    string left = clean.Substring(0, clean.IndexOf('=')).Trim();
                    if (seen.Add(left)) names.Add(left);
                }
            }
            catch { }
            return names;
        }

        // ==================== HTML/Text extractores ====================

        private string ToPlainText(string htmlOrText)
        {
            string s = htmlOrText ?? string.Empty;
            bool looksHtml = s.IndexOf('<') >= 0 && s.IndexOf('>') > s.IndexOf('<');

            if (looksHtml)
            {
                s = WebUtility.HtmlDecode(s);
                s = Regex.Replace(s, @"</?(?:span|div|p|i|b|strong|em|u|var|sub|sup|br)\b[^>]*>", m =>
                {
                    var tag = m.Value.ToLowerInvariant();
                    if (tag.StartsWith("<br") || tag.StartsWith("</p") || tag.StartsWith("<p"))
                        return "\n";
                    return string.Empty;
                }, RegexOptions.Singleline);
                s = Regex.Replace(s, "<[^>]+>", string.Empty, RegexOptions.Singleline);
            }
            else
            {
                s = WebUtility.HtmlDecode(s);
            }
            return NormalizeSpacesAndTokens(s);
        }

        private static string NormalizeSpacesAndTokens(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            // Espacios unicode → espacio normal
            s = s.Replace('\u2009', ' ')
                 .Replace('\u200A', ' ')
                 .Replace('\u202F', ' ')
                 .Replace('\u00A0', ' ')
                 .Replace('\u2002', ' ')
                 .Replace('\u2003', ' ')
                 .Replace('\u2005', ' ')
                 .Replace('\u2006', ' ');

            // '=' ancho completo → '=' ASCII
            s = s.Replace('\uFF1D', '=');

            s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            s = Regex.Replace(s, @"[ \t]+", " ");
            return s;
        }

        private Dictionary<string, string> ExtractResultBlocksByVar(string plainOutput, List<string> targetVars)
        {
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);
            if (string.IsNullOrEmpty(plainOutput) || targetVars == null || targetVars.Count == 0)
                return dict;

            var startRegex = new Regex(@"^\s*([A-Za-z_][A-Za-z0-9_'′,\.]*)\s*=", RegexOptions.Compiled);

            string currentVar = null;
            var sb = new StringBuilder();

            void Flush()
            {
                if (!string.IsNullOrEmpty(currentVar) && sb.Length > 0 && !dict.ContainsKey(currentVar))
                    dict[currentVar] = sb.ToString().Trim();
                sb.Clear();
            }

            foreach (var raw in plainOutput.Split(new[] { '\n' }, StringSplitOptions.None))
            {
                var line = raw.TrimEnd();

                var m = startRegex.Match(line);
                if (m.Success)
                {
                    string candidate = m.Groups[1].Value;
                    if (targetVars.Contains(candidate))
                    {
                        Flush();
                        currentVar = candidate;
                        sb.AppendLine(line.Trim());
                        continue;
                    }
                }

                if (!string.IsNullOrEmpty(currentVar))
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        sb.AppendLine(line);
                        continue;
                    }
                    sb.AppendLine(line.Trim());
                }
            }
            Flush();

            return dict;
        }

        private double ExtractFinalNumericFromBlock(string block)
        {
            if (string.IsNullOrWhiteSpace(block)) return double.NaN;

            try
            {
                int lastEq = block.LastIndexOf('=');
                string tail = lastEq >= 0 ? block.Substring(lastEq + 1) : block;
                tail = NormalizeSpacesAndTokens(tail).Trim();

                // Científica real con ×
                var sci = Regex.Match(
                    tail,
                    @"^\s*([+-]?\d+(?:\.\d+)?)\s*×\s*10\^?([+-]?\d+)\b",
                    RegexOptions.CultureInvariant
                );
                if (sci.Success)
                {
                    double a = double.Parse(sci.Groups[1].Value, CultureInfo.InvariantCulture);
                    int b = int.Parse(sci.Groups[2].Value, CultureInfo.InvariantCulture);
                    return a * Math.Pow(10, b);
                }

                // Decimal normal (anclado)
                var num = Regex.Match(tail, @"^\s*([+-]?\d+(?:\.\d+)?)\b", RegexOptions.CultureInvariant);
                if (num.Success)
                    return double.Parse(num.Groups[1].Value, CultureInfo.InvariantCulture);

                return double.NaN;
            }
            catch { return double.NaN; }
        }

        // Reemplaza el método ExtractFinalUnitFromBlock por esta versión (sin dígitos en unidad)
        private string ExtractFinalUnitFromBlock(string block)
        {
            if (string.IsNullOrWhiteSpace(block)) return string.Empty;

            try
            {
                int lastEq = block.LastIndexOf('=');
                string tail = lastEq >= 0 ? block.Substring(lastEq + 1) : block;
                tail = NormalizeSpacesAndTokens(tail).Trim();

                // número [× 10^exp] unidad(sin dígitos)
                var m = Regex.Match(
                    tail,
                    @"^\s*[+-]?\d+(?:\.\d+)?(?:\s*×\s*10\^?[+-]?\d+)?\s*(?<u>[" + UnitTokenClassNoDigits + @"]+)\b",
                    RegexOptions.CultureInvariant
                );
                if (m.Success)
                    return m.Groups["u"].Value.Trim();

                return string.Empty;
            }
            catch { return string.Empty; }
        }

        private Dictionary<string, string> ExtractResultBlocksByVarHtml(string html, List<string> targetVars)
        {
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);
            if (string.IsNullOrEmpty(html) || targetVars == null || targetVars.Count == 0)
                return dict;

            foreach (var fullName in targetVars)
            {
                SplitVar(fullName, out var baseVar, out var sub);
                string pat = sub == null
                    ? $@"<span\s+class=""eq"">\s*<var>\s*{Regex.Escape(baseVar)}\s*</var>[\s\S]*?</span>"
                    : $@"<span\s+class=""eq"">\s*<var>\s*{Regex.Escape(baseVar)}\s*</var>\s*<sub>\s*{Regex.Escape(sub)}\s*</sub>[\s\S]*?</span>";

                var m = Regex.Match(html, pat, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (m.Success)
                    dict[fullName] = m.Value;
            }
            return dict;
        }

        private static void SplitVar(string name, out string baseVar, out string sub)
        {
            var idx = name.IndexOf('_');
            if (idx > 0 && idx < name.Length - 1)
            {
                baseVar = name.Substring(0, idx);
                sub = name.Substring(idx + 1);
            }
            else
            {
                baseVar = name;
                sub = null;
            }
        }

        private string ExtractFinalUnitFromEqSpanHtml(string spanHtml)
        {
            if (string.IsNullOrWhiteSpace(spanHtml))
                return string.Empty;

            try
            {
                var eqTailMatch = Regex.Match(spanHtml, @"=(?!.*=)([\s\S]*)</span>", RegexOptions.Singleline);
                string tailHtml = eqTailMatch.Success ? eqTailMatch.Groups[1].Value : spanHtml;

                // Fracción con dvc/dvl
                var dvc = Regex.Match(tailHtml, @"<span\s+class=""dvc"">([\s\S]*?)</span>", RegexOptions.IgnoreCase);
                if (dvc.Success)
                {
                    string inner = dvc.Groups[1].Value;

                    var iMatches = Regex.Matches(inner, @"<i>(.*?)</i>", RegexOptions.Singleline);
                    string sup = "";
                    var supMatch = Regex.Match(inner, @"<sup>(.*?)</sup>", RegexOptions.Singleline);
                    if (supMatch.Success)
                        sup = "^" + WebUtility.HtmlDecode(supMatch.Groups[1].Value).Trim();

                    if (iMatches.Count >= 2)
                    {
                        string num = WebUtility.HtmlDecode(iMatches[0].Groups[1].Value).Trim();
                        string den = WebUtility.HtmlDecode(iMatches[1].Groups[1].Value).Trim();
                        return string.IsNullOrEmpty(sup) ? $"{num}/{den}" : $"{num}/{den}{sup}";
                    }
                    else if (iMatches.Count == 1)
                    {
                        string tok = WebUtility.HtmlDecode(iMatches[0].Groups[1].Value).Trim();
                        return string.IsNullOrEmpty(sup) ? tok : $"{tok}{sup}";
                    }
                }

                // Unidad simple: último <i>…</i> del tail
                var lastI = Regex.Matches(tailHtml, @"<i>(.*?)</i>", RegexOptions.Singleline);
                if (lastI.Count > 0)
                {
                    var u = WebUtility.HtmlDecode(lastI[lastI.Count - 1].Groups[1].Value).Trim();
                    return u;
                }
            }
            catch { /* ignore */ }

            return string.Empty;
        }

        // ==================== Utilidades ====================

        public string GetDebugInfo()
        {
            var info = new StringBuilder();
            info.AppendLine($"CPD Code ({_originalCode?.Length ?? 0} chars):");
            info.AppendLine(_originalCode ?? "No code");
            info.AppendLine();
            info.AppendLine($"Generated HTML/Text ({_lastHtmlResult?.Length ?? 0} chars):");
            info.AppendLine(_lastHtmlResult ?? "No HTML");
            return info.ToString();
        }

        public void Dispose()
        {
            try
            {
                _parser = null;
                _settings = null;
                _lastHtmlResult = null;
            }
            catch { }
        }

        private static string RemoveInlineComments(string rhs)
        {
            if (string.IsNullOrEmpty(rhs)) return rhs;

            rhs = NormalizeSpacesAndTokens(rhs).Trim();

            // Corta en #, ' (no "';'"), ’, ‘
            var m = Regex.Match(
                rhs,
                @"^(?<code>.*?)(?:\s*(?:#|'(?!;)|\u2019|\u2018).*)?$",
                RegexOptions.CultureInvariant
            );

            var code = m.Success ? m.Groups["code"].Value : rhs;
            return code.Trim();
        }

        private static string NormalizeSpaces(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Replace('\u2009', ' ')
                 .Replace('\u200A', ' ')
                 .Replace('\u202F', ' ')
                 .Replace('\u00A0', ' ')
                 .Replace('\u2002', ' ')
                 .Replace('\u2003', ' ')
                 .Replace('\u2005', ' ')
                 .Replace('\u2006', ' ');
            s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            s = Regex.Replace(s, @"[ \t]+", " ");
            return s;
        }

        // Reemplaza ExtractUnitSuffix por esta versión (sin dígitos en unidad)
        private static string ExtractUnitSuffix(string rhs)
        {
            if (string.IsNullOrWhiteSpace(rhs)) return string.Empty;
            rhs = NormalizeSpaces(rhs).Trim();

            // ?{...}[unidad]
            var mExp = Regex.Match(rhs,
                @"\?\s*\{[^}]*\}\s*(?<unit>[" + UnitTokenClassNoDigits + @"]+)?\s*$",
                RegexOptions.CultureInvariant);
            if (mExp.Success)
                return (mExp.Groups["unit"].Success ? mExp.Groups["unit"].Value : string.Empty).Trim();

            // número [unidad]
            var mNum = Regex.Match(rhs,
                @"[+-]?\d+(?:\.\d+)?\s*(?<unit>[" + UnitTokenClassNoDigits + @"]+)?\s*$",
                RegexOptions.CultureInvariant);
            if (mNum.Success)
                return (mNum.Groups["unit"].Success ? mNum.Groups["unit"].Value : string.Empty).Trim();

            return string.Empty;
        }

        // NUEVO (opcional): obtener los valores EXACTOS como texto, tal como aparecen en HTML (sin recorte)
        public List<string> GetResultValuesText()
        {
            var resultText = new List<string>();
            var vars = GetEquationVariableNamesFromCode();
            if (vars.Count == 0) return resultText;

            string plain = ToPlainText(_lastHtmlResult);
            var blocks = ExtractResultBlocksByVar(plain, vars);

            foreach (var v in vars)
            {
                if (blocks.TryGetValue(v, out var block))
                    resultText.Add(ExtractFinalNumericStringFromBlock(block));
                else
                    resultText.Add(string.Empty);
            }
            return resultText;
        }

        private string ExtractFinalNumericStringFromBlock(string block)
        {
            if (string.IsNullOrWhiteSpace(block)) return string.Empty;

            int lastEq = block.LastIndexOf('=');
            string tail = lastEq >= 0 ? block.Substring(lastEq + 1) : block;
            tail = NormalizeSpacesAndTokens(tail).Trim();

            // científica: base × 10^exp (con × real)
            var sci = Regex.Match(
                tail,
                @"^\s*([+-]?\d+(?:\.\d+)?)\s*×\s*10\^([+-]?\d+)\b",
                RegexOptions.CultureInvariant
            );
            if (sci.Success)
            {
                // Devolvemos exactamente el texto “a × 10^b”
                return $"{sci.Groups[1].Value} × 10^{sci.Groups[2].Value}";
            }

            // decimal normal: devolver todos los dígitos capturados
            var num = Regex.Match(tail, @"^\s*([+-]?\d+(?:\.\d+)?)\b", RegexOptions.CultureInvariant);
            if (num.Success)
                return num.Groups[1].Value;

            return string.Empty;
        }

        // Preprocesado de aliases de unidades no soportadas por el parser
        private string PreprocessCode(string code)
        {
            string s = code ?? string.Empty;
            return NormalizeUnsupportedUnits(s);
        }

        internal static string NormalizeUnsupportedUnits(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            // ton_f / tonf / tf → 1000 kgf (inserta espacio; no toca identificadores)
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])ton_f(?![A-Za-z0-9_])",
                " 1000 kgf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])tonf(?![A-Za-z0-9_])",
                " 1000 kgf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])tf(?![A-Za-z0-9_])",
                " 1000 kgf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );

            // Otras variantes: kip_f / kip / kips / klbf → 1000 lbf
            // s = Regex.Replace(s, @"(?<![A-Za-z0-9_])tonf(?![A-Za-z0-9_])", " 1000 kgf", RegexOptions.IgnoreCase);
            // s = Regex.Replace(s, @"(?<![A-Za-z0-9_])kip_f(?![A-Za-z0-9_])", " 1000 lbf", RegexOptions.IgnoreCase);
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])kip_f(?![A-Za-z0-9_])",
                " 1000 lbf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])kips?(?![A-Za-z0-9_])",
                " 1000 lbf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])kip(?![A-Za-z0-9_])",
                " 1000 lbf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
            s = Regex.Replace(
                s,
                @"(?<![A-Za-z0-9_])klbf(?![A-Za-z0-9_])",
                " 1000 lbf",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );

            s = Regex.Replace(s, @"[ \t]{2,}", " ");
            return s;
        }
    }
}
