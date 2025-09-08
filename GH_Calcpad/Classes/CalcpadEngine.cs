using PyCalcpad;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace GH_Calcpad.Classes
{
    /// <summary>
    /// Calcpad calculation engine optimized for Grasshopper.
    /// Designed for automated runs and optimization algorithms.
    /// </summary>
    public class CalcpadEngine : IDisposable
    {
        #region Properties

        /// <summary>
        /// Input variable names extracted from CPD code.
        /// </summary>
        public List<string> InputVariables { get; private set; }

        /// <summary>
        /// Current numeric values for the input variables.
        /// </summary>
        public List<double> InputValues { get; private set; }

        /// <summary>
        /// Output (result) variable names detected in equations.
        /// </summary>
        public List<string> OutputVariables { get; private set; }

        /// <summary>
        /// Current numeric values of the output variables.
        /// </summary>
        public List<double> OutputValues { get; private set; }

        /// <summary>
        /// True when the engine has code loaded and is ready to calculate.
        /// </summary>
        public bool IsReady => _calculator != null && !string.IsNullOrEmpty(_cleanCode);

        /// <summary>
        /// Duration of the last calculation (milliseconds).
        /// </summary>
        public double LastCalculationTime { get; private set; }

        #endregion

        #region Private Fields

        private Calculator _calculator;
        private Settings _settings;
        private string _originalCode;
        private string _cleanCode;
        private bool _disposed;

        #endregion

        #region Constructor & Initialization

        /// <summary>
        /// Creates a new CalcpadEngine instance and initializes internal calculator.
        /// </summary>
        public CalcpadEngine()
        {
            InputVariables = new List<string>();
            InputValues = new List<double>();
            OutputVariables = new List<string>();
            OutputValues = new List<double>();

            InitializeCalculator();
        }

        private void InitializeCalculator()
        {
            try
            {
                _settings = new Settings();
                _settings.Math.Decimals = 15;
                _calculator = new Calculator(_settings.Math);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to initialize PyCalcpad: {ex.Message}");
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Loads CPD source code and prepares variable tracking.
        /// </summary>
        /// <param name="cpdCode">Full CPD source text.</param>
        /// <param name="inputVariables">Input variable names.</param>
        /// <param name="inputValues">Initial values for the input variables.</param>
        public void LoadCode(string cpdCode, List<string> inputVariables, List<double> inputValues)
        {
            if (string.IsNullOrEmpty(cpdCode))
                throw new ArgumentException("CPD code cannot be empty");

            if (inputVariables.Count != inputValues.Count)
                throw new ArgumentException("Variables and values must have the same count");

            _originalCode = cpdCode;
            _cleanCode = CleanAndFixCode(cpdCode);

            InputVariables.Clear();
            InputVariables.AddRange(inputVariables);

            InputValues.Clear();
            InputValues.AddRange(inputValues);

            // Detect output variables (equations producing results).
            IdentifyOutputVariables();
        }

        /// <summary>
        /// Replaces current input variable values (fast path for iterative algorithms).
        /// </summary>
        public void UpdateInputValues(List<double> newValues)
        {
            if (newValues.Count != InputVariables.Count)
                throw new ArgumentException($"Expected {InputVariables.Count} values, received {newValues.Count}");

            InputValues.Clear();
            InputValues.AddRange(newValues);
        }

        /// <summary>
        /// Executes the calculation using the currently loaded code and input values.
        /// </summary>
        /// <returns>True if execution succeeded; false otherwise.</returns>
        public bool Calculate()
        {
            if (!IsReady)
                throw new InvalidOperationException("Engine is not ready. Use LoadCode() first.");

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                // Assign input variable values.
                for (int i = 0; i < InputVariables.Count; i++)
                {
                    if (IsValidVariableName(InputVariables[i]))
                    {
                        _calculator.SetVariable(InputVariables[i], InputValues[i]);
                    }
                }

                // Run calculation.
                _calculator.Run(_cleanCode);

                // Pull output values.
                ExtractOutputValues();

                stopwatch.Stop();
                LastCalculationTime = stopwatch.Elapsed.TotalMilliseconds;

                return true;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                LastCalculationTime = stopwatch.Elapsed.TotalMilliseconds;

                System.Diagnostics.Debug.WriteLine($"Calculation error: {ex.Message}");

                // Populate with NaN on failure.
                OutputValues.Clear();
                for (int i = 0; i < OutputVariables.Count; i++)
                    OutputValues.Add(double.NaN);

                return false;
            }
        }

        /// <summary>
        /// Evaluates and returns the current value of a single variable.
        /// </summary>
        public double GetVariableValue(string variableName)
        {
            if (!IsReady)
                return double.NaN;

            try
            {
                string result = _calculator.Eval(variableName);
                if (double.TryParse(result, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                    return value;
            }
            catch
            {
                // Ignore and return NaN.
            }

            return double.NaN;
        }

        #endregion

        #region Private Methods

        private string CleanAndFixCode(string code)
        {
            if (string.IsNullOrEmpty(code))
                return string.Empty;

            string cleaned = code;

            try
            {
                // Remove BOM if present.
                if (cleaned.StartsWith("\uFEFF"))
                    cleaned = cleaned.Substring(1);

                // Fix known problematic variable names.
                cleaned = cleaned.Replace("f'h", "fh")
                                 .Replace("i,j_k1", "ijk1")
                                 .Replace("e_d", "ed")
                                 .Replace("f_g", "fg");

                // Normalize line endings.
                cleaned = cleaned.Replace("\r\n", "\n").Replace("\r", "\n");

                // Ensure trailing newline.
                if (!cleaned.EndsWith("\n"))
                    cleaned += "\n";

                return cleaned;
            }
            catch
            {
                // On failure, return original text.
                return code;
            }
        }

        private void IdentifyOutputVariables()
        {
            OutputVariables.Clear();

            if (string.IsNullOrEmpty(_cleanCode))
                return;

            try
            {
                var lines = _cleanCode.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    var trimmed = line.Trim();

                    if (IsCalculationLine(trimmed))
                    {
                        string varName = ExtractVariableName(trimmed);
                        if (!string.IsNullOrEmpty(varName))
                            OutputVariables.Add(varName);
                    }
                }

                // Initialize output values (NaN placeholders).
                OutputValues.Clear();
                for (int i = 0; i < OutputVariables.Count; i++)
                    OutputValues.Add(double.NaN);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error identifying output variables: {ex.Message}");
            }
        }

        private bool IsCalculationLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                return false;

            if (line.StartsWith("'") || line.StartsWith("#") || line.StartsWith("\""))
                return false;

            var parts = line.Split('=');
            if (parts.Length < 2)
                return false;

            string rightSide = parts[1].Trim();

            // Consider it a calculation if the RHS includes operators/functions.
            return rightSide.Contains("+") || rightSide.Contains("-") || rightSide.Contains("*") ||
                   rightSide.Contains("/") || rightSide.Contains("(") || rightSide.Contains("sqrt") ||
                   rightSide.Contains("^");
        }

        private string ExtractVariableName(string equation)
        {
            if (string.IsNullOrWhiteSpace(equation))
                return string.Empty;

            var parts = equation.Split('=');
            if (parts.Length == 0)
                return string.Empty;

            return parts[0].Trim();
        }

        private bool IsValidVariableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            return Regex.IsMatch(name, @"^[a-zA-Z][a-zA-Z0-9_]*$");
        }

        private void ExtractOutputValues()
        {
            OutputValues.Clear();

            foreach (var variable in OutputVariables)
            {
                double value = GetVariableValue(variable);
                OutputValues.Add(value);
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Managed resources.
                    _calculator = null;
                    _settings = null;
                }

                _disposed = true;
            }
        }

        #endregion
    }
}