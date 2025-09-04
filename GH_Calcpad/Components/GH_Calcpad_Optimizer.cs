using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Revolutionary component that converts CPD files into automatically
    /// optimizable objective functions. Specifically designed for optimization
    /// algorithms like Galapagos, Octopus, and genetic algorithms.
    /// 
    /// KEY INNOVATION: Auto-detection of design variables and objectives.
    /// </summary>
    public class GH_Calcpad_Optimizer : GH_Component
    {
        private Dictionary<string, double> _cache;
        private List<OptimizationResult> _history;
        private DateTime _lastUpdate;

        public GH_Calcpad_Optimizer()
          : base(
                "Calcpad Optimizer",     // Name
                "CPOptim",              // Nickname
                "Converts CPD into optimizable objective function with intelligent auto-detection",
                "Calcpad",              // Category
                "4. Execution & Optimization"  // Optimizer
            )
        {
            _cache = new Dictionary<string, double>();
            _history = new List<OptimizationResult>();
            _lastUpdate = DateTime.Now;
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter(
                "Sheet Object", "SO",
                "CalcpadSheet from Load CPD",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "Design Variables", "DV",
                "Design variables to optimize (auto-detected if empty)",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "Variable Values", "VV",
                "Current values of design variables (from optimizer)",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Objective Names", "ON",
                "Names of objective variables to minimize/maximize",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Optimization Mode", "OM",
                "minimize, maximize, or target (per objective)",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "Target Values", "TV",
                "Target values for 'target' mode (optional)",
                GH_ParamAccess.list);

            // Make some parameters optional
            p[1].Optional = true; // Auto-detect variables
            p[4].Optional = true; // Default: minimize
            p[5].Optional = true; // Only for target mode
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddNumberParameter(
                "Fitness", "F",
                "Fitness value for optimizer (lower = better)",
                GH_ParamAccess.item);
            p.AddNumberParameter(
                "Objective Values", "OV",
                "Individual values of each objective",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Status", "ST",
                "Calculation status and convergence information",
                GH_ParamAccess.item);
            p.AddNumberParameter(
                "Iteration", "IT",
                "Current iteration number",
                GH_ParamAccess.item);
            p.AddNumberParameter(
                "Best Fitness", "BF",
                "Best fitness found so far",
                GH_ParamAccess.item);
            p.AddTextParameter(
                "Suggested Variables", "SV",
                "Auto-detected variables for optimization",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Convergence Info", "CI",
                "Convergence and progress information",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Get inputs
            object data = null;
            var designVars = new List<string>();
            var varValues = new List<double>();
            var objectiveNames = new List<string>();
            var optimModes = new List<string>();
            var targetValues = new List<double>();

            if (!DA.GetData(0, ref data)) return;
            DA.GetDataList(1, designVars);
            DA.GetDataList(2, varValues);
            DA.GetDataList(3, objectiveNames);
            DA.GetDataList(4, optimModes);
            DA.GetDataList(5, targetValues);

            // 2) Unwrap CalcpadSheet
            CalcpadSheet sheet = ExtractSheet(data);
            if (sheet == null) return;

            try
            {
                // 3) Auto-detect design variables if not specified
                if (designVars.Count == 0)
                {
                    designVars = AutoDetectDesignVariables(sheet);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                        $"🤖 Auto-detected {designVars.Count} design variables");
                }

                // 4) Auto-detect objectives if not specified
                if (objectiveNames.Count == 0)
                {
                    objectiveNames = AutoDetectObjectives(sheet);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                        $"🎯 Auto-detected {objectiveNames.Count} objectives");
                }

                // 5) Validate consistency
                if (varValues.Count != designVars.Count)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        $"Variable Values ({varValues.Count}) must match Design Variables ({designVars.Count})");
                    return;
                }

                // 6) Create cache key
                string cacheKey = CreateCacheKey(varValues);
                
                // 7) Check cache to avoid recalculations
                if (_cache.ContainsKey(cacheKey))
                {
                    double cachedFitness = _cache[cacheKey];
                    OutputCachedResult(DA, cachedFitness, designVars, objectiveNames);
                    return;
                }

                // 8) Execute calculation with new variables
                var fitness = EvaluateObjectives(sheet, designVars, varValues, objectiveNames, optimModes, targetValues, DA);
                
                // 9) Save to cache and history
                _cache[cacheKey] = fitness.totalFitness;
                _history.Add(new OptimizationResult
                {
                    Iteration = _history.Count + 1,
                    Variables = new List<double>(varValues),
                    Fitness = fitness.totalFitness,
                    Objectives = fitness.objectiveValues,
                    Timestamp = DateTime.Now
                });

                // 10) Convergence analysis
                var convergenceInfo = AnalyzeConvergence();

                // 11) Set outputs
                DA.SetData(0, fitness.totalFitness);
                DA.SetDataList(1, fitness.objectiveValues);
                DA.SetData(2, fitness.status);
                DA.SetData(3, _history.Count);
                DA.SetData(4, _history.Count > 0 ? _history.Min(h => h.Fitness) : fitness.totalFitness);
                DA.SetDataList(5, designVars);
                DA.SetData(6, convergenceInfo);

            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Optimization error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts CalcpadSheet from input object
        /// </summary>
        private CalcpadSheet ExtractSheet(object data)
        {
            if (data is GH_ObjectWrapper wrapper)
            {
                var sheet = wrapper.Value as CalcpadSheet;
                if (sheet == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
                        "The received object is not a valid CalcpadSheet.");
                }
                return sheet;
            }
            else
            {
                var sheet = data as CalcpadSheet;
                if (sheet == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, 
                        "The received object is not a valid CalcpadSheet.");
                }
                return sheet;
            }
        }

        /// <summary>
        /// Auto-detects design variables by analyzing CPD code patterns
        /// </summary>
        private List<string> AutoDetectDesignVariables(CalcpadSheet sheet)
        {
            var candidates = new List<string>();
            
            // Strategy 1: Variables with "round" values (probably parameters)
            for (int i = 0; i < sheet.Variables.Count; i++)
            {
                var value = sheet.Values[i];
                var unit = sheet.Units[i];
                
                // Variables with "round" values are candidates
                if (IsRoundValue(value) && !string.IsNullOrEmpty(unit))
                {
                    candidates.Add(sheet.Variables[i]);
                }
            }

            // Strategy 2: Variables with typical design dimension units
            var designUnits = new[] { "mm", "m", "cm", "kN", "MPa", "kg", "°", "rad" };
            for (int i = 0; i < sheet.Variables.Count; i++)
            {
                var unit = sheet.Units[i].ToLower();
                if (designUnits.Any(u => unit.Contains(u)))
                {
                    candidates.Add(sheet.Variables[i]);
                }
            }

            return candidates.Distinct().Take(10).ToList(); // Limit to 10 variables
        }

        /// <summary>
        /// Checks if a value is "round" (probably a design parameter)
        /// </summary>
        private bool IsRoundValue(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return false;

            // Values that are multiples of 5, 10, 25, 50, etc.
            var roundFactors = new[] { 1, 5, 10, 25, 50, 100, 250, 500, 1000 };
            
            foreach (var factor in roundFactors)
            {
                if (Math.Abs(value % factor) < 0.001)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Auto-detects objectives by analyzing result equations
        /// </summary>
        private List<string> AutoDetectObjectives(CalcpadSheet sheet)
        {
            var objectives = new List<string>();
            
            // Search for variables representing typical results
            var objectiveKeywords = new[] { "stress", "weight", "cost", "area", "volume", "force", "moment", "deflection" };
            
            foreach (var variable in sheet.Variables)
            {
                var lowerVar = variable.ToLower();
                if (objectiveKeywords.Any(keyword => lowerVar.Contains(keyword)))
                {
                    objectives.Add(variable);
                }
            }

            // If no obvious objectives found, use last calculated variables
            if (objectives.Count == 0 && sheet.Variables.Count > 0)
            {
                objectives.Add(sheet.Variables.Last());
            }

            return objectives.Take(3).ToList(); // Limit to 3 objectives
        }

        /// <summary>
        /// Creates cache key based on variable values
        /// </summary>
        private string CreateCacheKey(List<double> values)
        {
            var sb = new StringBuilder();
            foreach (var value in values)
            {
                sb.Append(value.ToString("F6"));
                sb.Append("_");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Outputs result from cache
        /// </summary>
        private void OutputCachedResult(IGH_DataAccess DA, double cachedFitness, List<string> designVars, List<string> objectiveNames)
        {
            // Search for result in history
            var lastResult = _history.LastOrDefault(h => Math.Abs(h.Fitness - cachedFitness) < 1e-6);
            
            if (lastResult != null)
            {
                DA.SetData(0, lastResult.Fitness);
                DA.SetDataList(1, lastResult.Objectives);
                DA.SetData(2, "✅ Result from cache");
                DA.SetData(3, lastResult.Iteration);
                DA.SetData(4, _history.Min(h => h.Fitness));
                DA.SetDataList(5, designVars);
                DA.SetData(6, "Cache hit - no recalculation");
            }
            else
            {
                // Fallback if not found in history
                DA.SetData(0, cachedFitness);
                DA.SetDataList(1, new List<double>());
                DA.SetData(2, "✅ Cache");
                DA.SetData(3, _history.Count);
                DA.SetData(4, cachedFitness);
                DA.SetDataList(5, designVars);
                DA.SetData(6, "Basic cache");
            }
        }

        /// <summary>
        /// Evaluates objective functions with current values
        /// </summary>
        private (double totalFitness, List<double> objectiveValues, string status) EvaluateObjectives(
            CalcpadSheet sheet, List<string> designVars, List<double> varValues,
            List<string> objectiveNames, List<string> optimModes, List<double> targetValues, IGH_DataAccess DA)
        {
            try
            {
                // Apply new variable values
                for (int i = 0; i < designVars.Count; i++)
                {
                    sheet.SetVariable(designVars[i], varValues[i]);
                }

                // Execute calculation
                sheet.Calculate();

                // Get results
                var equations = sheet.GetResultEquations();
                var results = sheet.GetResultValues();

                // Extract objective values
                var objectiveValues = new List<double>();
                for (int i = 0; i < objectiveNames.Count; i++)
                {
                    var objName = objectiveNames[i];
                    var objIndex = equations.FindIndex(eq => ExtractVariableName(eq) == objName);
                    
                    if (objIndex >= 0 && objIndex < results.Count)
                    {
                        objectiveValues.Add(results[objIndex]);
                    }
                    else
                    {
                        objectiveValues.Add(double.MaxValue); // Penalize objectives not found
                    }
                }

                // Calculate total fitness
                double totalFitness = CalculateTotalFitness(objectiveValues, optimModes, targetValues);
                
                string status = $"✅ Successful calculation | Objectives: {objectiveValues.Count} | Fitness: {totalFitness:F6}";
                
                return (totalFitness, objectiveValues, status);
            }
            catch (Exception ex)
            {
                var errorValues = objectiveNames.Select(_ => double.MaxValue).ToList();
                return (double.MaxValue, errorValues, $"❌ Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts variable name from equation
        /// </summary>
        private string ExtractVariableName(string equation)
        {
            if (string.IsNullOrWhiteSpace(equation))
                return string.Empty;

            int equalIndex = equation.IndexOf('=');
            if (equalIndex <= 0)
                return string.Empty;

            return equation.Substring(0, equalIndex).Trim();
        }

        /// <summary>
        /// Calculates total fitness combining multiple objectives
        /// </summary>
        private double CalculateTotalFitness(List<double> objectiveValues, List<string> optimModes, List<double> targetValues)
        {
            if (objectiveValues.Count == 0)
                return double.MaxValue;

            double totalFitness = 0.0;
            
            for (int i = 0; i < objectiveValues.Count; i++)
            {
                double objValue = objectiveValues[i];
                string mode = i < optimModes.Count ? optimModes[i].ToLower() : "minimize";
                
                double fitness;
                
                switch (mode)
                {
                    case "maximize":
                        fitness = -objValue; // Invert to maximize
                        break;
                    case "target":
                        double target = i < targetValues.Count ? targetValues[i] : 0.0;
                        fitness = Math.Abs(objValue - target); // Minimize distance to target
                        break;
                    case "minimize":
                    default:
                        fitness = objValue;
                        break;
                }
                
                totalFitness += fitness;
            }
            
            return totalFitness;
        }

        /// <summary>
        /// Analyzes optimization convergence
        /// </summary>
        private string AnalyzeConvergence()
        {
            if (_history.Count < 2)
                return "Starting optimization...";

            // ✅ FIXED: Use Skip() instead of TakeLast() for .NET Framework 4.8
            int recentCount = Math.Min(10, _history.Count);
            var recentResults = _history.Skip(_history.Count - recentCount).ToList();
            var fitnessValues = recentResults.Select(r => r.Fitness).ToList();
            
            // Calculate trend
            double improvement = fitnessValues.First() - fitnessValues.Last();
            double improvementRate = improvement / Math.Max(fitnessValues.First(), 1e-6) * 100;
            
            // Detect stagnation - using Skip() instead of TakeLast()
            int lastFiveCount = Math.Min(5, fitnessValues.Count);
            var lastFiveValues = fitnessValues.Skip(fitnessValues.Count - lastFiveCount).ToList();
            bool isStagnant = lastFiveValues.All(f => Math.Abs(f - fitnessValues.Last()) < 1e-6);
            
            var sb = new StringBuilder();
            sb.AppendLine($"Iteration {_history.Count}:");
            sb.AppendLine($"Improvement: {improvement:F6} ({improvementRate:F2}%)");
            sb.AppendLine($"Best fitness: {_history.Min(h => h.Fitness):F6}");
            
            if (isStagnant)
                sb.AppendLine("⚠️ Possible stagnation detected");
            else if (improvementRate > 1)
                sb.AppendLine("📈 Significant improvement");
            else
                sb.AppendLine("🔄 Converging...");
                
            return sb.ToString();
        }

        public override Guid ComponentGuid
            => new Guid("F7A8B9C0-D1E2-4F3A-5B6C-7D8E9F0A1B2C");

        protected override Bitmap Icon
            => Resources.Icon_Next;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }

    /// <summary>
    /// Helper class to store optimization results
    /// </summary>
    public class OptimizationResult
    {
        public int Iteration { get; set; }
        public List<double> Variables { get; set; }
        public double Fitness { get; set; }
        public List<double> Objectives { get; set; }
        public DateTime Timestamp { get; set; }
    }
}