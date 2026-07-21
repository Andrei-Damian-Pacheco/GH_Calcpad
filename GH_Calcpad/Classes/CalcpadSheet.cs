using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace GH_Calcpad.Classes
{
    public class CalcpadSheet
    {
        // Convención de nombres: variables marcadas con estos prefijos dentro del .cpd
        // se filtran directamente por nombre en vez de depender de heurísticas
        // (nombres redondos, palabras clave en inglés, etc.). Archivos que no usan
        // ningún prefijo siguen funcionando igual que siempre (arrays posicionales).
        public const string DesignVariablePrefix = "gh_";
        public const string ObjectiveVariablePrefix = "ghc_";

        public List<string> Variables { get; }
        public List<double> Values { get; }
        public List<string> Units { get; }

        private string _originalCode;

        private const string UnitTokenClassNoDigits = "A-Za-z°µμΩ℧·/\\-\\^²³";

        // Resultado crudo del worker JSON (Fase 1 de la migración): nombre -> (valor, unidad)
        // para cada línea de asignación que Calcpad calculó al resolver la hoja.
        private Dictionary<string, (double Value, string Unit)> _workerResults;

        public string OriginalCode => _originalCode ?? string.Empty;
        public bool HasCodeAvailable => !string.IsNullOrEmpty(_originalCode);
        public string CodeInfo => string.IsNullOrEmpty(_originalCode) ? "No CPD code" : $"CPD code: {_originalCode.Length} characters";

        public CalcpadSheet(List<string> variables, List<double> values, List<string> units)
        {
            Variables = variables ?? new List<string>();
            Values = values ?? new List<double>();
            Units = units ?? new List<string>();
            _originalCode = string.Empty;
        }

        /// <summary>
        /// Independent copy of this sheet (own Variables/Values/Units lists and own code string).
        /// Grasshopper fans a single output out to multiple inputs by sharing the same object
        /// reference, not by duplicating it - so any component that mutates a CalcpadSheet
        /// (SetVariable/SetUnit) must clone it first, or two branches reading the same
        /// Load CPD output would silently overwrite each other's values.
        /// </summary>
        public CalcpadSheet Clone()
        {
            var clone = new CalcpadSheet(new List<string>(Variables), new List<double>(Values), new List<string>(Units));
            clone.SetFullCode(_originalCode);
            return clone;
        }

        public void SetFullCode(string code)
        {
            _originalCode = code ?? string.Empty;
        }

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
                // End-of-line lookahead is "\r?\n|\z", not bare "$": in Multiline mode .NET's $
                // matches only right before a literal \n, not before \r\n - so on a Windows
                // (CRLF) .cpd file the lazy/greedy rhs group would get stuck right before the
                // \r with nowhere left to go, and the whole match would silently fail (no
                // exception - SetVariable would just leave the line untouched).
                string pattern = @"(?m)(^|\s*';'\s*)\s*(?<lhs>" + Regex.Escape(name) + @")\s*=\s*(?<rhs>[^\r\n]*?)(?=(\s*';'\s*|\r?\n|\z))";

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
                });

                if (!replaced)
                {
                    string linePattern = @"(?m)^(?<pre>\s*)" + Regex.Escape(name) + @"\s*=\s*(?<rhs>[^\r\n]+)(?=\r?\n|\z)";
                    _originalCode = Regex.Replace(_originalCode, linePattern, m =>
                    {
                        string pre = m.Groups["pre"].Value;
                        string rhs = m.Groups["rhs"].Value;
                        string unit = ExtractUnitSuffix(rhs);
                        string newRhs = string.IsNullOrEmpty(unit) ? vStr : (vStr + " " + unit);
                        return $"{pre}{name} = {newRhs}";
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetVariable('{name}', {value}) - ERROR: {ex.Message}");
            }
        }

        public void Calculate()
        {
            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code to calculate. Use SetFullCode() first.");

            string codeToParse = PreprocessCode(_originalCode);
            CalcpadWorkerResult workerResult;
            try
            {
                workerResult = CalcpadWorkerClient.Instance.Solve(codeToParse);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Calcpad worker unavailable: {ex.Message}");
            }

            if (!workerResult.Ok)
                throw new InvalidOperationException($"Calculation error: {workerResult.Error}");

            _workerResults = workerResult.Results;
        }

        // ==================== Resultados ====================

        /// <summary>
        /// Every worker-computed assignment (literal design variable or computed
        /// equation alike - the worker doesn't distinguish, both are "assignments")
        /// whose name starts with <paramref name="prefix"/>. Reads directly from the
        /// last Calculate() worker response, so it reflects whatever value was
        /// actually just used, regardless of whether it came from Play's positional
        /// Values input or its named Design Names/Values input.
        /// </summary>
        public void GetResultsByPrefix(string prefix, out List<string> names, out List<double> values, out List<string> units)
        {
            names = new List<string>();
            values = new List<double>();
            units = new List<string>();
            if (_workerResults == null || string.IsNullOrEmpty(prefix)) return;

            foreach (var kv in _workerResults)
            {
                if (kv.Key.StartsWith(prefix, StringComparison.Ordinal))
                {
                    names.Add(kv.Key);
                    values.Add(kv.Value.Value);
                    units.Add(kv.Value.Unit ?? string.Empty);
                }
            }
        }

        /// <summary>
        /// Exact (case-insensitive) lookup of a single computed assignment by name, straight
        /// from the worker's authoritative result set - no source-text scanning involved.
        /// Returns the canonical (correctly-cased) name alongside the value/unit, since a
        /// caller-typed name may not match the .cpd file's exact casing.
        /// </summary>
        public bool TryGetResult(string name, out string canonicalName, out double value, out string unit)
        {
            canonicalName = null;
            value = double.NaN;
            unit = string.Empty;
            if (_workerResults == null || string.IsNullOrEmpty(name)) return false;

            foreach (var kv in _workerResults)
            {
                if (string.Equals(kv.Key, name, StringComparison.OrdinalIgnoreCase))
                {
                    canonicalName = kv.Key;
                    value = kv.Value.Value;
                    unit = kv.Value.Unit ?? string.Empty;
                    return true;
                }
            }
            return false;
        }

        public List<string> GetResultEquations()
        {
            var equations = new List<string>();
            try
            {
                var lines = _originalCode.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    var clean = line.Trim();
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

        // ==================== Detección de ecuaciones ====================

        private bool IsEquationDefinition(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            int firstEq = line.IndexOf('=');
            if (firstEq < 0) return false;
            if (line.Contains("';'")) return false;
            if (line.IndexOf('=', firstEq + 1) >= 0) return false;

            string left = line.Substring(0, firstEq).Trim();
            string rightRaw = line.Substring(firstEq + 1).Trim();
            string right = RemoveInlineComments(rightRaw);

            if (!Regex.IsMatch(left, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$")) return false;

            string unitClass = CalcpadSyntax.Instance.UnitCharClass;
            var rxNumUnit = new Regex(@"^[+-]?\d+(?:\.\d+)?(?:\s*[" + unitClass + @"]+)?\s*$", RegexOptions.CultureInvariant);
            if (rxNumUnit.IsMatch(right)) return false;

            var rxExplicit = new Regex(@"^\?\s*\{\s*[^}]+\s*\}\s*(?:[" + unitClass + @"]+)?\s*$", RegexOptions.CultureInvariant);
            if (rxExplicit.IsMatch(right)) return false;

            bool hasOps = right.IndexOfAny(new[] { '+', '-', '*', '/', '^', '(', ')' }) >= 0
                       || Regex.IsMatch(right, @"\b(sqrt|sin|cos|tan|log|exp|abs|min|max|pow)\b", RegexOptions.IgnoreCase);

            bool isVariableReference = Regex.IsMatch(right, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$", RegexOptions.CultureInvariant);
            return hasOps || isVariableReference;
        }

        // ==================== Utilidades varias ====================

        private static string RemoveInlineComments(string rhs)
        {
            if (string.IsNullOrEmpty(rhs)) return rhs;
            rhs = NormalizeSpaces(rhs).Trim();
            var m = Regex.Match(rhs, @"^(?<code>.*?)(?:\s*(?:#|'(?!;)|\u2019|\u2018).*)?$");
            var code = m.Success ? m.Groups["code"].Value : rhs;
            return code.Trim();
        }

        private static string NormalizeSpaces(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Replace('\u2009', ' ').Replace('\u200A', ' ').Replace('\u202F', ' ')
                 .Replace('\u00A0', ' ').Replace('\u2002', ' ').Replace('\u2003', ' ')
                 .Replace('\u2005', ' ').Replace('\u2006', ' ');
            s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            s = Regex.Replace(s, @"[ \t]+", " ");
            return s;
        }

        private static string ExtractUnitSuffix(string rhs)
        {
            if (string.IsNullOrWhiteSpace(rhs)) return string.Empty;
            rhs = NormalizeSpaces(rhs).Trim();

            var mExp = Regex.Match(rhs, @"\?\s*\{[^}]*\}\s*(?<unit>[" + UnitTokenClassNoDigits + @"]+)?\s*$");
            if (mExp.Success)
                return (mExp.Groups["unit"].Success ? mExp.Groups["unit"].Value : string.Empty).Trim();

            var mNum = Regex.Match(rhs, @"[+-]?\d+(?:\.\d+)?\s*(?<unit>[" + UnitTokenClassNoDigits + @"]+)?\s*$");
            if (mNum.Success)
                return (mNum.Groups["unit"].Success ? mNum.Groups["unit"].Value : string.Empty).Trim();

            return string.Empty;
        }

        private string PreprocessCode(string code) => NormalizeUnsupportedUnits(code ?? string.Empty);

        internal static string NormalizeUnsupportedUnits(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])ton_f(?![A-Za-z0-9_])", " 1000 kgf", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])tonf(?![A-Za-z0-9_])", " 1000 kgf", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])tf(?![A-Za-z0-9_])", " 1000 kgf", RegexOptions.IgnoreCase);

            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])kip_f(?![A-Za-z0-9_])", " 1000 lbf", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])kips?(?![A-Za-z0-9_])", " 1000 lbf", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])kip(?![A-Za-z0-9_])", " 1000 lbf", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(?<![A-Za-z0-9_])klbf(?![A-Za-z0-9_])", " 1000 lbf", RegexOptions.IgnoreCase);

            s = Regex.Replace(s, @"[ \t]{2,}", " ");
            return s;
        }

    }
}
