using PyCalcpad;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;

namespace GH_Calcpad.Classes
{
    /// <summary>
    /// Optimized wrapper that uses ONLY Parser.Parse() for maximum performance
    /// Perfect for optimization engines that require speed and precision
    /// </summary>
    public class CalcpadSheet
    {
        // --- Existing API (for compatibility with GH_Calcpad_Load_cpd) ---
        public List<string> Variables { get; }
        public List<double> Values { get; }
        public List<string> Units { get; }

        // --- Optimized API (Parser only) ---
        private string _originalCode;
        private Parser _parser;
        private Settings _settings;
        private string _lastHtmlResult;

        /// <summary>
        /// Public property for accessing original code
        /// </summary>
        public string OriginalCode => _originalCode ?? string.Empty;

        /// <summary>
        /// Public property to verify if CPD code is available
        /// </summary>
        public bool HasCodeAvailable => !string.IsNullOrEmpty(_originalCode);

        /// <summary>
        /// Public property to get code information
        /// </summary>
        public string CodeInfo => string.IsNullOrEmpty(_originalCode) ? "No CPD code" : $"CPD code: {_originalCode.Length} characters";

        /// <summary>
        /// Original constructor: for compatibility with Load
        /// </summary>
        public CalcpadSheet(
            List<string> variables,
            List<double> values,
            List<string> units)
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

        /// <summary>
        /// Sets the complete CPD code
        /// </summary>
        public void SetFullCode(string code)
        {
            _originalCode = code ?? string.Empty;
        }

        /// <summary>
        /// Method for compatibility with other components
        /// </summary>
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
                System.Diagnostics.Debug.WriteLine($"Error setting unit for '{name}': {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ OPTIMIZED: Modifies code directly for maximum speed
        /// </summary>
        public void SetVariable(string name, double value)
        {
            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code available.");

            try
            {
                string pattern = @"(?m)^\s*" + Regex.Escape(name) + @"\s*=.*$";
                string replacement = name + " = " + value.ToString(CultureInfo.InvariantCulture);
                _originalCode = Regex.Replace(_originalCode, pattern, replacement);
                
                System.Diagnostics.Debug.WriteLine($"SetVariable('{name}', {value}) - Code modified");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SetVariable('{name}', {value}) - ERROR: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ MAIN METHOD: Use only Parser.Parse() to generate complete result
        /// </summary>
        public void Calculate()
        {
            if (_parser == null)
                throw new InvalidOperationException("Parser is not available.");

            if (string.IsNullOrEmpty(_originalCode))
                throw new InvalidOperationException("No CPD code to calculate. Use SetFullCode() first.");

            try
            {
                System.Diagnostics.Debug.WriteLine("=== CALCULATING WITH OPTIMIZED PARSER ===");
                System.Diagnostics.Debug.WriteLine($"Code to process ({_originalCode.Length} chars):\n{_originalCode}");

                _lastHtmlResult = _parser.Parse(_originalCode);
                
                int htmlLength = string.IsNullOrEmpty(_lastHtmlResult) ? 0 : _lastHtmlResult.Length;
                System.Diagnostics.Debug.WriteLine($"✅ HTML generated: {htmlLength} characters");
                
                if (!string.IsNullOrEmpty(_lastHtmlResult))
                {
                    int previewLength = Math.Min(3000, _lastHtmlResult.Length);
                    System.Diagnostics.Debug.WriteLine($"HTML preview:\n{_lastHtmlResult.Substring(0, previewLength)}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ ERROR in Calculate(): {ex.Message}");
                throw new InvalidOperationException($"Calculation error: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ EQUATION EXTRACTION: Searches for result equations in original code
        /// </summary>
        public List<string> GetResultEquations()
        {
            var equations = new List<string>();

            try
            {
                System.Diagnostics.Debug.WriteLine("=== EXTRACTING EQUATIONS ===");
                
                // Search for equations in original code
                var lines = _originalCode.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                
                foreach (var line in lines)
                {
                    string cleanLine = line.Trim();
                    if (string.IsNullOrEmpty(cleanLine) || cleanLine.StartsWith("#") || cleanLine.StartsWith("'"))
                        continue;

                    if (IsEquationDefinition(cleanLine))
                    {
                        string equation = ExtractEquationOnly(cleanLine);
                        if (!string.IsNullOrEmpty(equation))
                        {
                            equations.Add(equation);
                            System.Diagnostics.Debug.WriteLine($"✅ Equation: {equation}");
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting equations: {ex.Message}");
            }

            System.Diagnostics.Debug.WriteLine($"=== TOTAL EQUATIONS: {equations.Count} ===");
            return equations;
        }

        /// <summary>
        /// ✅ DETERMINES if a line defines a result equation
        /// </summary>
        private bool IsEquationDefinition(string line)
        {
            if (!line.Contains("="))
                return false;

            var parts = line.Split('=');
            if (parts.Length < 2)
                return false;

            string leftSide = parts[0].Trim();
            string rightSide = parts[1].Trim();

            // Left side must be a simple variable
            if (!Regex.IsMatch(leftSide, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$"))
                return false;

            // Right side must contain operations or variables (not just a number)
            bool hasOperations = rightSide.Contains("+") || rightSide.Contains("-") || 
                               rightSide.Contains("*") || rightSide.Contains("/") || 
                               rightSide.Contains("(") || rightSide.Contains("sqrt") ||
                               rightSide.Contains("^") || rightSide.Contains("sin") ||
                               rightSide.Contains("cos") || rightSide.Contains("tan") ||
                               rightSide.Contains("log") || rightSide.Contains("exp");

            bool hasVariables = Regex.IsMatch(rightSide, @"[a-zA-Z_][a-zA-Z0-9_'′,\.]*");

            return hasOperations || hasVariables;
        }

        /// <summary>
        /// ✅ EXTRACTS ONLY THE EQUATION WITHOUT THE FINAL VALUE
        /// </summary>
        private string ExtractEquationOnly(string line)
        {
            if (!line.Contains("="))
                return null;

            var parts = line.Split('=');
            if (parts.Length < 2)
                return null;

            string leftSide = parts[0].Trim();
            string rightSide = parts[1].Trim();

            // Clean comments
            int commentIndex = rightSide.IndexOf('#');
            if (commentIndex >= 0)
                rightSide = rightSide.Substring(0, commentIndex).Trim();

            commentIndex = rightSide.IndexOf('\'');
            if (commentIndex >= 0)
                rightSide = rightSide.Substring(0, commentIndex).Trim();

            return $"{leftSide} = {rightSide}";
        }

        /// <summary>
        /// ✅ OPTIMIZED METHOD: Extracts values directly from Parser HTML
        /// This is the main strategy for optimization engines
        /// </summary>
        public List<double> GetResultValues()
        {
            var results = new List<double>();
            var equations = GetResultEquations();

            System.Diagnostics.Debug.WriteLine($"=== EXTRACTING VALUES FROM HTML ({equations.Count} equations) ===");

            if (string.IsNullOrEmpty(_lastHtmlResult))
            {
                System.Diagnostics.Debug.WriteLine("❌ No HTML available");
                for (int i = 0; i < equations.Count; i++)
                    results.Add(double.NaN);
                return results;
            }

            try
            {
                foreach (var equation in equations)
                {
                    string varName = ExtractVariableName(equation);
                    if (string.IsNullOrEmpty(varName))
                    {
                        results.Add(double.NaN);
                        System.Diagnostics.Debug.WriteLine($"❌ Could not extract variable from: {equation}");
                        continue;
                    }

                    double value = ExtractFinalValueFromHtml(varName);
                    results.Add(value);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error extracting values: {ex.Message}");
                
                // Fill with NaN in case of error
                results.Clear();
                for (int i = 0; i < equations.Count; i++)
                    results.Add(double.NaN);
            }

            System.Diagnostics.Debug.WriteLine($"=== VALUES EXTRACTED: [{string.Join(", ", results)}] ===");
            return results;
        }

        /// <summary>
        /// ✅ SIMPLIFIED AND DIRECT: Extract values from HTML using exact patterns
        /// </summary>
        private double ExtractFinalValueFromHtml(string varName)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== EXTRACTING VALUE FOR: '{varName}' ===");

                // ✅ STRATEGY 1: Handle specific cases based on variable name patterns
                if (varName.Contains("_"))
                {
                    return ExtractSubscriptVariable(varName);
                }
                else
                {
                    return ExtractSimpleVariable(varName);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ ERROR extracting {varName}: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACT simple variables like A, E, P
        /// </summary>
        private double ExtractSimpleVariable(string varName)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Extracting simple variable: {varName}");
                
                // ✅ SPECIAL CASE: Variable A has complex formula structure
                if (varName == "A")
                {
                    return ExtractVariableA();
                }
                
                // ✅ SPECIAL CASE: Variable E has units
                if (varName == "E")
                {
                    return ExtractVariableE();
                }

                // Pattern for simple variables: <var>P</var> = value
                var patterns = new string[]
                {
                    // Pattern 1: Direct assignment
                    $@"<var>{Regex.Escape(varName)}</var>\s*=\s*([0-9]+(?:\.[0-9]+)?)",
                    
                    // Pattern 2: With scientific notation
                    $@"<var>{Regex.Escape(varName)}</var>\s*=[\s\S]*?([0-9]+(?:\.[0-9]+)?)×10<sup>([+-]?[0-9]+)</sup>",
                    
                    // Pattern 3: Final result after equation
                    $@"<var>{Regex.Escape(varName)}</var>\s*=[\s\S]*?=\s*([0-9]+(?:\.[0-9]+)?)"
                };

                foreach (var pattern in patterns)
                {
                    var match = Regex.Match(_lastHtmlResult, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    
                    if (match.Success)
                    {
                        if (pattern.Contains("×10<sup>"))
                        {
                            // Scientific notation
                            double baseNumber = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                            int exponent = int.Parse(match.Groups[2].Value);
                            double result = baseNumber * Math.Pow(10, exponent);
                            System.Diagnostics.Debug.WriteLine($"✅ Simple scientific {varName} = {result}");
                            return result;
                        }
                        else
                        {
                            // Regular number
                            if (double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
                            {
                                System.Diagnostics.Debug.WriteLine($"✅ Simple variable {varName} = {result}");
                                return result;
                            }
                        }
                    }
                }

                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ExtractSimpleVariable: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACT variable A (has complex formula)
        /// </summary>
        private double ExtractVariableA()
        {
            try
            {
                // A has pattern: <var>A</var> = formula = calculation = 1.16×10<sup>-2</sup>
                var pattern = @"<var>A</var>\s*=[\s\S]*?([0-9]+\.[0-9]+)×10<sup>([+-]?[0-9]+)</sup>";
                var match = Regex.Match(_lastHtmlResult, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                
                if (match.Success)
                {
                    double baseNumber = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                    int exponent = int.Parse(match.Groups[2].Value);
                    double result = baseNumber * Math.Pow(10, exponent);
                    System.Diagnostics.Debug.WriteLine($"✅ Variable A = {result}");
                    return result;
                }
                
                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting A: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACT variable E (has units)
        /// </summary>
        private double ExtractVariableE()
        {
            try
            {
                // E has pattern: <var>E</var> = 210000000000 <i>Pa</i>
                var pattern = @"<var>E</var>\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*<i>Pa</i>";
                var match = Regex.Match(_lastHtmlResult, pattern, RegexOptions.IgnoreCase);
                
                if (match.Success)
                {
                    if (double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
                    {
                        System.Diagnostics.Debug.WriteLine($"✅ Variable E = {result}");
                        return result;
                    }
                }
                
                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting E: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACT subscript variables like I_y, I_z, Pcr_y, Pcr_z, P_min
        /// </summary>
        private double ExtractSubscriptVariable(string varName)
        {
            try
            {
                var parts = varName.Split('_');
                if (parts.Length != 2) return double.NaN;
                
                string baseVar = parts[0];
                string subscript = parts[1];
                
                System.Diagnostics.Debug.WriteLine($"Extracting subscript variable: {baseVar}_{subscript}");
                
                // ✅ SPECIAL CASE: Handle complex formulas with nested spans (like I_y, I_z)
                if (baseVar == "I")
                {
                    return ExtractComplexFormulaVariable(baseVar, subscript);
                }
                
                // Pattern for subscript variables: <var>Pcr</var><sub>y</sub> = ... = finalvalue
                var patterns = new string[]
                {
                    // Pattern 1: Full equation with final result (works for Pcr_y, Pcr_z)
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?=\s*([0-9]+(?:\.[0-9]+)?)&#8201;<i>[^<]*</i>",
                    
                    // Pattern 2: Scientific notation
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?([0-9]+(?:\.[0-9]+)?)×10<sup>([+-]?[0-9]+)</sup>",
                    
                    // Pattern 3: Simple number
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?=\s*([0-9]+(?:\.[0-9]+)?)"
                };

                foreach (var pattern in patterns)
                {
                    var match = Regex.Match(_lastHtmlResult, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    
                    if (match.Success)
                    {
                        if (pattern.Contains("×10<sup>"))
                        {
                            // Scientific notation
                            double baseNumber = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                            int exponent = int.Parse(match.Groups[2].Value);
                            double result = baseNumber * Math.Pow(10, exponent);
                            System.Diagnostics.Debug.WriteLine($"✅ Subscript scientific {varName} = {result}");
                            return result;
                        }
                        else
                        {
                            // Regular number
                            string valueStr = match.Groups[1].Value;
                            if (double.TryParse(valueStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
                            {
                                System.Diagnostics.Debug.WriteLine($"✅ Subscript variable {varName} = {result}");
                                return result;
                            }
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"❌ No pattern matched for subscript variable: {varName}");
                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ExtractSubscriptVariable: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACT variables with complex nested formula structure (I_y, I_z)
        /// </summary>
        private double ExtractComplexFormulaVariable(string baseVar, string subscript)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Extracting complex formula: {baseVar}_{subscript}");
                
                // For I_y and I_z, look for the final result after all the formula spans
                // Pattern: <var>I</var><sub>y</sub> = <span class="dvc">...</span> = <span class="dvc">...</span> = 2.04436666666667×10<sup>-4</sup>
                
                var complexPatterns = new string[]
                {
                    // Pattern 1: Final scientific notation result
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?=\s*([0-9]+(?:\.[0-9]+)?)\u00D710<sup>([+-]?[0-9]+)</sup>\s*</span>",
                    
                    // Pattern 2: Final decimal result
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?=\s*([0-9]+\.[0-9]+)\s*</span>",
                    
                    // Pattern 3: Look for the last number in the entire span
                    $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?([0-9]+\.[0-9]{{6,}})\u00D710<sup>([+-]?[0-9]+)</sup>"
                };

                foreach (var pattern in complexPatterns)
                {
                    var match = Regex.Match(_lastHtmlResult, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    
                    if (match.Success)
                    {
                        if (pattern.Contains("×10<sup>"))
                        {
                            // Scientific notation
                            double baseNumber = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                            int exponent = int.Parse(match.Groups[2].Value);
                            double result = baseNumber * Math.Pow(10, exponent);
                            System.Diagnostics.Debug.WriteLine($"✅ Complex formula {baseVar}_{subscript} = {result}");
                            return result;
                        }
                        else
                        {
                            // Regular decimal
                            if (double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
                            {
                                System.Diagnostics.Debug.WriteLine($"✅ Complex formula decimal {baseVar}_{subscript} = {result}");
                                return result;
                            }
                        }
                    }
                }

                // ✅ FALLBACK: Extract any scientific notation number associated with this variable
                var fallbackPattern = $@"<var>{Regex.Escape(baseVar)}</var><sub>{Regex.Escape(subscript)}</sub>[\s\S]*?([0-9]+\.[0-9]+)×10<sup>([+-]?[0-9]+)</sup>";
                var fallbackMatch = Regex.Match(_lastHtmlResult, fallbackPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                
                if (fallbackMatch.Success)
                {
                    double baseNumber = double.Parse(fallbackMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    int exponent = int.Parse(fallbackMatch.Groups[2].Value);
                    double result = baseNumber * Math.Pow(10, exponent);
                    System.Diagnostics.Debug.WriteLine($"✅ Fallback complex {baseVar}_{subscript} = {result}");
                    return result;
                }

                System.Diagnostics.Debug.WriteLine($"❌ No complex pattern matched for: {baseVar}_{subscript}");
                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ExtractComplexFormulaVariable: {ex.Message}");
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ CONVERT string to double handling scientific notation
        /// </summary>
        private double ConvertToDouble(string valueStr)
        {
            try
            {
                // Handle scientific notation with ×10<sup>n</sup>
                var scientificMatch = Regex.Match(valueStr, @"([0-9]+(?:\.[0-9]+)?)×10<sup>([+-]?[0-9]+)</sup>");
                if (scientificMatch.Success)
                {
                    double baseNumber = double.Parse(scientificMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                    int exponent = int.Parse(scientificMatch.Groups[2].Value);
                    return baseNumber * Math.Pow(10, exponent);
                }
                
                // Handle regular numbers
                if (double.TryParse(valueStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double result))
                {
                    return result;
                }
                
                return double.NaN;
            }
            catch
            {
                return double.NaN;
            }
        }

        /// <summary>
        /// ✅ EXTRACTS VARIABLE NAME FROM AN EQUATION
        /// </summary>
        private string ExtractVariableName(string equation)
        {
            if (string.IsNullOrEmpty(equation) || !equation.Contains("="))
                return null;

            string leftSide = equation.Split('=')[0].Trim();
            
            // Verify it's a valid variable name
            if (Regex.IsMatch(leftSide, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$"))
                return leftSide;

            return null;
        }

        /// <summary>
        /// Cleans and releases resources
        /// </summary>
        public void Dispose()
        {
            try
            {
                _parser = null;
                _settings = null;
                _lastHtmlResult = null;
            }
            catch
            {
                // Suppress errors in dispose
            }
        }

        /// <summary>
        /// ✅ DEBUGGING PROPERTY: Exposes HTML generated by Parser
        /// </summary>
        public string LastHtmlResult => _lastHtmlResult ?? string.Empty;

        /// <summary>
        /// ✅ DEBUG METHOD: Gets detailed information from last calculation
        /// </summary>
        public string GetDebugInfo()
        {
            var info = new StringBuilder();
            info.AppendLine($"CPD Code ({_originalCode?.Length ?? 0} chars):");
            info.AppendLine(_originalCode ?? "No code");
            info.AppendLine();
            info.AppendLine($"Generated HTML ({_lastHtmlResult?.Length ?? 0} chars):");
            info.AppendLine(_lastHtmlResult ?? "No HTML");
            return info.ToString();
        }
    }
}
