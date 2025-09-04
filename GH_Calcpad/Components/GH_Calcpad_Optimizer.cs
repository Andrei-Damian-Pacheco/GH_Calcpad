using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Converts CPD into optimizable objective functions.
    /// Auto-detects design variables and objectives when not provided.
    /// </summary>
    public class GH_Calcpad_Optimizer : GH_Component
    {
        private Dictionary<string, OptimizationResult> _cache;
        private List<OptimizationResult> _history;
        private DateTime _lastUpdate;
        private string _lastSignature;

        public GH_Calcpad_Optimizer()
          : base(
                "Calcpad Optimizer",
                "CPOptim",
                "Converts CPD into optimizable objective function with intelligent auto-detection",
                "Calcpad",
                "4. Execution & Optimization"
            )
        {
            _cache = new Dictionary<string, OptimizationResult>(StringComparer.Ordinal);
            _history = new List<OptimizationResult>();
            _lastUpdate = DateTime.Now;
            _lastSignature = string.Empty;
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Sheet Object", "SO", "CalcpadSheet from Load CPD", GH_ParamAccess.item);
            p.AddTextParameter("Design Variables", "DV", "Design variables to optimize (auto-detected if empty)", GH_ParamAccess.list);
            p.AddNumberParameter("Variable Values", "VV", "Current values of design variables (from optimizer)", GH_ParamAccess.list);
            p.AddTextParameter("Objective Names", "ON", "Names of objective variables to minimize/maximize", GH_ParamAccess.list);
            p.AddTextParameter("Optimization Mode", "OM", "minimize, maximize, or target (per objective)", GH_ParamAccess.list);
            p.AddNumberParameter("Target Values", "TV", "Target values for 'target' mode (optional)", GH_ParamAccess.list);

            p[1].Optional = true; // Auto-detect variables
            p[4].Optional = true; // Default: minimize
            p[5].Optional = true; // Only for target mode
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddNumberParameter("Fitness", "F", "Fitness value for optimizer (lower = better)", GH_ParamAccess.item);
            p.AddNumberParameter("Objective Values", "OV", "Individual values of each objective", GH_ParamAccess.list);
            p.AddTextParameter("Status", "ST", "Calculation status and convergence information", GH_ParamAccess.item);
            p.AddNumberParameter("Iteration", "IT", "Current iteration number", GH_ParamAccess.item);
            p.AddNumberParameter("Best Fitness", "BF", "Best fitness found so far", GH_ParamAccess.item);
            p.AddTextParameter("Suggested Variables", "SV", "Auto-detected variables for optimization", GH_ParamAccess.list);
            p.AddTextParameter("Convergence Info", "CI", "Convergence and progress information", GH_ParamAccess.item);
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
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Auto-detected {designVars.Count} design variables");
                }

                // 4) Auto-detect objectives if not specified
                if (objectiveNames.Count == 0)
                {
                    objectiveNames = AutoDetectObjectives(sheet);
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Auto-detected {objectiveNames.Count} objectives");
                }

                // 5) Validate consistency
                if (varValues.Count != designVars.Count)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        $"Variable Values ({varValues.Count}) must match Design Variables ({designVars.Count})");
                    return;
                }

                // 6) Normalize inputs (trim names, lower modes)
                designVars = designVars.Select(s => (s ?? string.Empty).Trim()).Where(s => s.Length > 0).ToList();
                objectiveNames = objectiveNames.Select(s => (s ?? string.Empty).Trim()).Where(s => s.Length > 0).ToList();
                optimModes = optimModes.Select(s => (s ?? "minimize").Trim()).ToList();

                // 7) Reset cache/history if problem signature changed
                string signature = BuildSignature(designVars, objectiveNames, optimModes, targetValues);
                MaybeResetCache(signature);

                // 8) Create cache key (includes DVs, values, objectives, modes, targets)
                string cacheKey = CreateCacheKey(designVars, varValues, objectiveNames, optimModes, targetValues);

                // 9) Check cache
                if (_cache.TryGetValue(cacheKey, out var cached))
                {
                    OutputCachedResult(DA, cached, designVars);
                    return;
                }

                // 10) Execute calculation with new variables
                var eval = EvaluateObjectives(sheet, designVars, varValues, objectiveNames, optimModes, targetValues, DA);

                // 11) Save to cache and history
                var result = new OptimizationResult
                {
                    Iteration = _history.Count + 1,
                    Variables = new List<double>(varValues),
                    Fitness = eval.totalFitness,
                    Objectives = eval.objectiveValues,
                    Timestamp = DateTime.Now
                };
                _cache[cacheKey] = result;
                _history.Add(result);

                // 12) Convergence analysis
                var convergenceInfo = AnalyzeConvergence();

                // 13) Outputs
                DA.SetData(0, result.Fitness);
                DA.SetDataList(1, result.Objectives);
                DA.SetData(2, eval.status);
                DA.SetData(3, result.Iteration);
                DA.SetData(4, _history.Count > 0 ? _history.Min(h => h.Fitness) : result.Fitness);
                DA.SetDataList(5, designVars);
                DA.SetData(6, convergenceInfo);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Optimization error: {ex.Message}");
            }
        }

        private CalcpadSheet ExtractSheet(object data)
        {
            if (data is GH_ObjectWrapper wrapper)
            {
                var sheet = wrapper.Value as CalcpadSheet;
                if (sheet == null)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return sheet;
            }
            else
            {
                var sheet = data as CalcpadSheet;
                if (sheet == null)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return sheet;
            }
        }

        // Preferir variables provistas por el loader; si no hay, heurística
        private List<string> AutoDetectDesignVariables(CalcpadSheet sheet)
        {
            var result = new List<string>();

            // Si el loader aportó variables (explícitas o normales), úsalas directamente
            if (sheet.Variables != null && sheet.Variables.Count > 0)
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var v in sheet.Variables)
                {
                    if (seen.Add(v)) result.Add(v);
                    if (result.Count >= 10) break;
                }
                return result;
            }

            // Fallback heurístico
            var candidates = new List<string>();
            for (int i = 0; i < sheet.Variables.Count; i++)
            {
                var value = (i < sheet.Values.Count) ? sheet.Values[i] : double.NaN;
                var unit = (i < sheet.Units.Count) ? sheet.Units[i] : string.Empty;

                if (IsRoundValue(value) && !string.IsNullOrEmpty(unit))
                    candidates.Add(sheet.Variables[i]);
            }

            var designUnits = new[] { "mm", "m", "cm", "kn", "mpa", "kg", "°", "rad" }; // minúsculas
            for (int i = 0; i < sheet.Variables.Count; i++)
            {
                var unit = (i < sheet.Units.Count ? sheet.Units[i] : string.Empty).ToLowerInvariant();
                if (designUnits.Any(u => unit.Contains(u)))
                    candidates.Add(sheet.Variables[i]);
            }

            return candidates.Distinct().Take(10).ToList();
        }

        private bool IsRoundValue(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return false;
            var roundFactors = new[] { 1, 5, 10, 25, 50, 100, 250, 500, 1000 };
            foreach (var factor in roundFactors)
            {
                if (Math.Abs(value % factor) < 1e-3)
                    return true;
            }
            return false;
        }

        private List<string> AutoDetectObjectives(CalcpadSheet sheet)
        {
            var objectives = new List<string>();

            // Variables de ecuaciones reales (no asignaciones simples)
            var eqs = sheet.GetResultEquations();
            var lhs = eqs.Select(ExtractVariableName).Where(s => !string.IsNullOrEmpty(s)).ToList();

            var objectiveKeywords = new[] { "stress", "weight", "cost", "area", "volume", "force", "moment", "deflection" };

            foreach (var name in lhs)
            {
                var lower = name.ToLowerInvariant();
                if (objectiveKeywords.Any(k => lower.Contains(k)))
                    objectives.Add(name);
            }

            if (objectives.Count == 0 && lhs.Count > 0)
            {
                // Usa los últimos resultados calculados si no hay match por palabra clave
                objectives.AddRange(lhs.Skip(Math.Max(0, lhs.Count - 3)));
            }

            return objectives.Take(3).ToList();
        }

        private string BuildSignature(List<string> dv, List<string> on, List<string> om, List<double> tv)
        {
            var sb = new StringBuilder();
            sb.Append("DV:");
            sb.Append(string.Join("|", dv));
            sb.Append(";ON:");
            sb.Append(string.Join("|", on));
            sb.Append(";OM:");
            sb.Append(string.Join("|", om.Select(m => (m ?? "minimize").ToLowerInvariant())));
            sb.Append(";TV:");
            sb.Append(string.Join("|", tv.Select(v => v.ToString("R", CultureInfo.InvariantCulture))));
            return sb.ToString();
        }

        private void MaybeResetCache(string signature)
        {
            if (!string.Equals(_lastSignature, signature, StringComparison.Ordinal))
            {
                _cache.Clear();
                _history.Clear();
                _lastSignature = signature;
            }
        }

        private string CreateCacheKey(
            List<string> designVars, List<double> values,
            List<string> objectiveNames, List<string> optimModes, List<double> targetValues)
        {
            var sb = new StringBuilder();

            sb.Append("DV:");
            for (int i = 0; i < designVars.Count; i++)
            {
                sb.Append(designVars[i]);
                sb.Append('=');
                double v = i < values.Count ? values[i] : double.NaN;
                sb.Append(v.ToString("R", CultureInfo.InvariantCulture));
                sb.Append(';');
            }

            sb.Append("|ON:");
            sb.Append(string.Join(",", objectiveNames));

            sb.Append("|OM:");
            sb.Append(string.Join(",", optimModes.Select(m => (m ?? "minimize").ToLowerInvariant())));

            sb.Append("|TV:");
            sb.Append(string.Join(",", targetValues.Select(t => t.ToString("R", CultureInfo.InvariantCulture))));

            return sb.ToString();
        }

        private void OutputCachedResult(IGH_DataAccess DA, OptimizationResult cached, List<string> designVars)
        {
            DA.SetData(0, cached.Fitness);
            DA.SetDataList(1, cached.Objectives ?? new List<double>());
            DA.SetData(2, "✅ Result from cache");
            DA.SetData(3, cached.Iteration);
            DA.SetData(4, _history.Count > 0 ? _history.Min(h => h.Fitness) : cached.Fitness);
            DA.SetDataList(5, designVars);
            DA.SetData(6, "Cache hit - no recalculation");
        }

        private (double totalFitness, List<double> objectiveValues, string status) EvaluateObjectives(
            CalcpadSheet sheet, List<string> designVars, List<double> varValues,
            List<string> objectiveNames, List<string> optimModes, List<double> targetValues, IGH_DataAccess DA)
        {
            try
            {
                // Aplicar valores nuevos
                for (int i = 0; i < designVars.Count; i++)
                    sheet.SetVariable(designVars[i], varValues[i]);

                // Ejecutar cálculo
                sheet.Calculate();

                // Extraer resultados
                var equations = sheet.GetResultEquations();           // "LHS = RHS"
                var results = sheet.GetResultValues();                // valores numéricos
                var lhs = equations.Select(ExtractVariableName).ToList();

                // Mapa nombre->índice para acceso O(1)
                var indexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < lhs.Count; i++)
                {
                    if (!indexByName.ContainsKey(lhs[i]))
                        indexByName[lhs[i]] = i;
                }

                // Valores de objetivos en el mismo orden solicitado
                var objectiveValues = new List<double>(objectiveNames.Count);
                foreach (var objName in objectiveNames)
                {
                    if (indexByName.TryGetValue(objName, out int idx) && idx >= 0 && idx < results.Count)
                        objectiveValues.Add(results[idx]);
                    else
                        objectiveValues.Add(1e30); // Penalización si no se encuentra
                }

                // Fitness total
                double totalFitness = CalculateTotalFitness(objectiveValues, optimModes, targetValues);

                string status = $"✅ Successful calculation | Objectives: {objectiveValues.Count} | Fitness: {totalFitness:G6}";
                return (totalFitness, objectiveValues, status);
            }
            catch (Exception ex)
            {
                var errorValues = Enumerable.Repeat(1e30, Math.Max(1, objectiveNames.Count)).ToList();
                return (1e30, errorValues, $"❌ Error: {ex.Message}");
            }
        }

        private string ExtractVariableName(string equation)
        {
            if (string.IsNullOrWhiteSpace(equation)) return string.Empty;
            int equalIndex = equation.IndexOf('=');
            if (equalIndex <= 0) return string.Empty;
            return equation.Substring(0, equalIndex).Trim();
        }

        private double CalculateTotalFitness(List<double> objectiveValues, List<string> optimModes, List<double> targetValues)
        {
            if (objectiveValues == null || objectiveValues.Count == 0)
                return 1e30;

            double total = 0.0;
            for (int i = 0; i < objectiveValues.Count; i++)
            {
                double val = objectiveValues[i];
                if (double.IsNaN(val) || double.IsInfinity(val)) val = 1e30;

                string mode = (i < optimModes.Count ? optimModes[i] : "minimize")?.ToLowerInvariant() ?? "minimize";
                double fitness;

                switch (mode)
                {
                    case "maximize":
                        fitness = -val; // invertir para maximizar
                        break;
                    case "target":
                        double target = i < targetValues.Count ? targetValues[i] : 0.0;
                        fitness = Math.Abs(val - target);
                        break;
                    case "minimize":
                    default:
                        fitness = val;
                        break;
                }

                total += fitness;
            }
            return total;
        }

        private string AnalyzeConvergence()
        {
            if (_history.Count < 2)
                return "Starting optimization...";

            int recentCount = Math.Min(10, _history.Count);
            var recentResults = _history.Skip(_history.Count - recentCount).ToList();
            var fitnessValues = recentResults.Select(r => r.Fitness).ToList();

            double improvement = fitnessValues.First() - fitnessValues.Last();
            double improvementRate = improvement / Math.Max(Math.Abs(fitnessValues.First()), 1e-6) * 100.0;

            int lastFiveCount = Math.Min(5, fitnessValues.Count);
            var lastFiveValues = fitnessValues.Skip(fitnessValues.Count - lastFiveCount).ToList();
            bool isStagnant = lastFiveValues.All(f => Math.Abs(f - fitnessValues.Last()) < 1e-6);

            var sb = new StringBuilder();
            sb.AppendLine($"Iteration {_history.Count}:");
            sb.AppendLine($"Improvement: {improvement:G6} ({improvementRate:F2}%)");
            sb.AppendLine($"Best fitness: {_history.Min(h => h.Fitness):G6}");

            if (isStagnant) sb.AppendLine("⚠️ Possible stagnation detected");
            else if (improvementRate > 1) sb.AppendLine("📈 Significant improvement");
            else sb.AppendLine("🔄 Converging...");

            return sb.ToString();
        }

        public override Guid ComponentGuid => new Guid("F7A8B9C0-D1E2-4F3A-5B6C-7D8E9F0A1B2C");
        protected override Bitmap Icon => Resources.Icon_Next;
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }

    public class OptimizationResult
    {
        public int Iteration { get; set; }
        public List<double> Variables { get; set; }
        public double Fitness { get; set; }
        public List<double> Objectives { get; set; }
        public DateTime Timestamp { get; set; }
    }
}