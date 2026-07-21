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
    /// Sets specific variables on a CalcpadSheet by name before it reaches Play CPD.
    /// Workflow: Load CPD -> Search Variables -> Play CPD
    /// </summary>
    public class GH_Calcpad_Search_Variables : GH_Component
    {
        public GH_Calcpad_Search_Variables()
          : base(
                "Search Variables",
                "SearchVars",
                "Sets specific variables on a CalcpadSheet by name (defaults to all 'gh_' design variables if Filter Names is left empty)",
                "Calcpad",
                "3. Variable Modification"
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "CalcpadSheet instance (from Load CPD)", GH_ParamAccess.item);
            p.AddTextParameter(
                "Filter Names", "FN",
                $"Names of variables to modify. Leave empty/unconnected to auto-select every variable prefixed '{CalcpadSheet.DesignVariablePrefix}'.",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "New Values", "NV",
                "New values, 1:1 order with Filter Names (or with the auto-selected 'gh_' variables, in file order, if Filter Names is empty).",
                GH_ParamAccess.list);
            p[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "Independent copy of the input Sheet with the requested variables set - connect to Play CPD", GH_ParamAccess.item);
            p.AddTextParameter("Modified Names", "MN", "Names of variables that were modified successfully", GH_ParamAccess.list);
            p.AddNumberParameter("Modified Values", "MV", "Corresponding modified values", GH_ParamAccess.list);
            p.AddTextParameter("Not Found", "NF", "Requested variables that were not found in the Sheet", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object data = null;
            var filterNamesIn = new List<string>();
            var newValues = new List<double>();

            if (!DA.GetData(0, ref data)) return;
            DA.GetDataList(1, filterNamesIn);
            if (!DA.GetDataList(2, newValues)) return;

            CalcpadSheet incomingSheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (incomingSheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return;
            }

            // Clone before mutating: the same upstream Sheet could be wired into more than one
            // Search Variables/Play CPD (e.g. comparing scenarios side by side), and CalcpadSheet
            // is a reference type - Grasshopper fans a wire out by sharing the object, not copying it.
            CalcpadSheet sheet = incomingSheet.Clone();

            List<string> filterNames = filterNamesIn.Count > 0
                ? filterNamesIn
                : GetDesignVariableNames(sheet);

            if (filterNames.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                    filterNamesIn.Count > 0
                        ? "Filter Names is empty."
                        : $"No variables prefixed '{CalcpadSheet.DesignVariablePrefix}' were found, and Filter Names was left empty.");
                DA.SetData(0, new GH_ObjectWrapper(sheet));
                DA.SetDataList(1, new List<string>());
                DA.SetDataList(2, new List<double>());
                DA.SetDataList(3, new List<string>());
                return;
            }

            if (newValues.Count != filterNames.Count)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    $"Filter Names ({filterNames.Count}) and New Values ({newValues.Count}) must have the same length.");
            }

            // Case-insensitive lookup for matching the user's typed/auto-detected name against the
            // canonical, correctly-cased name Load CPD detected - SetVariable rewrites the .cpd code
            // with an exact-case regex match, so the canonical name (not the user's) must be used.
            var canonicalByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in sheet.Variables)
                canonicalByName[v] = v;

            var modifiedNames = new List<string>();
            var modifiedValues = new List<double>();
            var notFoundNames = new List<string>();

            int n = Math.Min(filterNames.Count, newValues.Count);
            for (int i = 0; i < n; i++)
            {
                var requestedName = filterNames[i]?.Trim();
                if (string.IsNullOrEmpty(requestedName))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Empty name found in Filter Names at position {i}.");
                    continue;
                }

                if (!canonicalByName.TryGetValue(requestedName, out var canonicalName))
                {
                    notFoundNames.Add(requestedName);
                    continue;
                }

                double value = newValues[i];
                try
                {
                    sheet.SetVariable(canonicalName, value);
                    modifiedNames.Add(canonicalName);
                    modifiedValues.Add(value);
                }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Could not assign '{canonicalName}': {ex.Message}");
                }
            }

            if (modifiedNames.Count > 0)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Modified {modifiedNames.Count} of {filterNames.Count} requested variables.");
            if (notFoundNames.Count > 0)
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"{notFoundNames.Count} variables not found: {string.Join(", ", notFoundNames)}.");

            DA.SetData(0, new GH_ObjectWrapper(sheet));
            DA.SetDataList(1, modifiedNames);
            DA.SetDataList(2, modifiedValues);
            DA.SetDataList(3, notFoundNames);
        }

        private static List<string> GetDesignVariableNames(CalcpadSheet sheet)
        {
            var result = new List<string>();
            foreach (var name in sheet.Variables)
                if (name.StartsWith(CalcpadSheet.DesignVariablePrefix, StringComparison.Ordinal))
                    result.Add(name);
            return result;
        }

        public override Guid ComponentGuid
            => new Guid("A1F07C3D-4B8F-4E92-AB6C-DEADBEEF1234");

        protected override Bitmap Icon
            => Resources.Icon_SearchV;

        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}
