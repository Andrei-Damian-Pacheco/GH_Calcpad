using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Reads specific computed results from a calculated CalcpadSheet by name (defaults to
    /// all 'ghc_' objective variables if Filter Names is left empty). Mirrors Search
    /// Variables' Sheet + optional Filter Names pattern, but for reading results after Play
    /// CPD instead of setting inputs before it.
    /// </summary>
    public class GH_Calcpad_Search_Results : GH_Component
    {
        public GH_Calcpad_Search_Results()
          : base(
                "Search Results",
                "SearchRes",
                "Reads specific computed results from a calculated CalcpadSheet (defaults to all 'ghc_' objective variables if Filter Names is left empty)",
                "Calcpad",
                "5. Result Filtering"
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "Calculated CalcpadSheet (from Play CPD)", GH_ParamAccess.item);
            p.AddTextParameter(
                "Filter Names", "FN",
                $"Names of results to read. Leave empty/unconnected to auto-select every variable prefixed '{CalcpadSheet.ObjectiveVariablePrefix}'.",
                GH_ParamAccess.list);
            p[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("Equations", "EQ", "'Name = expression' for each match (falls back to 'Name = value unit' if no source formula is found, e.g. plain inputs)", GH_ParamAccess.list);
            p.AddNumberParameter("Values", "V", "Computed values 1:1 with Equations", GH_ParamAccess.list);
            p.AddTextParameter("Units", "U", "Units 1:1 with Equations", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object data = null;
            var filterNamesIn = new List<string>();

            if (!DA.GetData(0, ref data)) return;
            DA.GetDataList(1, filterNamesIn);

            CalcpadSheet sheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return;
            }

            List<string> requestedNames;
            if (filterNamesIn.Count > 0)
            {
                requestedNames = filterNamesIn;
            }
            else
            {
                sheet.GetResultsByPrefix(CalcpadSheet.ObjectiveVariablePrefix, out requestedNames, out _, out _);
                if (requestedNames.Count == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        $"No results prefixed '{CalcpadSheet.ObjectiveVariablePrefix}' were found, and Filter Names was left empty.");
                    DA.SetDataList(0, new List<string>());
                    DA.SetDataList(1, new List<double>());
                    DA.SetDataList(2, new List<string>());
                    return;
                }
            }

            var equationTextByName = BuildEquationTextByName(sheet);

            var equations = new List<string>();
            var values = new List<double>();
            var units = new List<string>();
            var notFoundNames = new List<string>();

            foreach (var requestedName in requestedNames)
            {
                var name = requestedName?.Trim();
                if (string.IsNullOrEmpty(name)) continue;

                if (!sheet.TryGetResult(name, out var canonicalName, out var value, out var unit))
                {
                    notFoundNames.Add(name);
                    continue;
                }

                string equationText = equationTextByName.TryGetValue(canonicalName, out var eq)
                    ? eq
                    : string.IsNullOrEmpty(unit) ? $"{canonicalName} = {value}" : $"{canonicalName} = {value} {unit}";

                equations.Add(equationText);
                values.Add(value);
                units.Add(unit);
            }

            if (notFoundNames.Count > 0)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"{notFoundNames.Count} results not found: {string.Join(", ", notFoundNames)}.");

            DA.SetDataList(0, equations);
            DA.SetDataList(1, values);
            DA.SetDataList(2, units);
        }

        // "Name = RHS" text for every source-code line that looks like a computed equation,
        // keyed by exact (canonical) name - lets us show the real formula instead of just the
        // number when one exists. Last occurrence wins if a name is assigned more than once,
        // matching the worker's own "final value" semantics.
        private static Dictionary<string, string> BuildEquationTextByName(CalcpadSheet sheet)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var eq in sheet.GetResultEquations())
            {
                int i = eq.IndexOf('=');
                if (i <= 0) continue;
                var lhs = eq.Substring(0, i).Trim();
                map[lhs] = eq;
            }
            return map;
        }

        public override Guid ComponentGuid
            => new Guid("D4E5F6A7-B8C9-4D0E-1F2A-3B4C5D6E7F8A");

        protected override Bitmap Icon
            => Resources.Icon_SearchR;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
