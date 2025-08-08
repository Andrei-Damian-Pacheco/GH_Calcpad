using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component for filtering and modifying specific variables while maintaining 
    /// complete structure for perfect integration with GH_Calcpad_Play.
    /// Workflow: Load → Search (modify specific values) → Play → Export
    /// </summary>
    public class GH_Calcpad_Search_Variables : GH_Component
    {
        public GH_Calcpad_Search_Variables()
          : base(
                "Search Variables",     // More descriptive name
                "SearchVars",                // Clearer nickname
                "Filters and modifies specific variables maintaining complete structure for Play CPD",
                "Calcpad",                // Category
                "3. Variable Modification"   // Modify Variables (Search)
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter(
                "All Names", "AN",
                "Complete list of variable names (from Load CPD)",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "All Values", "AV", 
                "Complete list of variable values (from Load CPD)",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Filter Names", "FN",
                "Names of specific variables to modify",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "New Values", "NV",
                "New values for filtered variables (1:1 order with Filter Names)",
                GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddNumberParameter(
                "All Values", "AV",
                "Complete array of values with modifications applied → connect to Play CPD",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Modified Names", "MN",
                "Names of variables that were modified successfully",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "Modified Values", "MV", 
                "Corresponding modified values (for visualization/debug)",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Not Found", "NF",
                "Requested variables that were not found",
                GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var allNames = new List<string>();
            var allValues = new List<double>();
            var filterNames = new List<string>();
            var newValues = new List<double>();

            // Validate required inputs
            if (!DA.GetDataList(0, allNames)) return;
            if (!DA.GetDataList(1, allValues)) return;
            if (!DA.GetDataList(2, filterNames)) return;
            if (!DA.GetDataList(3, newValues)) return;

            // Validate input data consistency
            if (allNames.Count != allValues.Count)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Error,
                    $"All Names ({allNames.Count}) and All Values ({allValues.Count}) must have the same length.");
                return;
            }

            if (filterNames.Count != newValues.Count)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Error,
                    $"Filter Names ({filterNames.Count}) and New Values ({newValues.Count}) must have the same length.");
                return;
            }

            // Validate there is data to process
            if (allNames.Count == 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    "No variables received in All Names.");
                DA.SetDataList(0, new List<double>());
                DA.SetDataList(1, new List<string>());
                DA.SetDataList(2, new List<double>());
                DA.SetDataList(3, filterNames);
                return;
            }

            // Create copy of all values (this will be our final result)
            var resultAllValues = new List<double>(allValues);

            // Create dictionary for fast name → index mapping
            var nameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < allNames.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(allNames[i]))
                {
                    var cleanName = allNames[i].Trim();
                    // If duplicates exist, keep the last index
                    nameToIndex[cleanName] = i;
                }
            }

            var modifiedNames = new List<string>();
            var modifiedValues = new List<double>();
            var notFoundNames = new List<string>();

            // Process each variable to modify
            for (int i = 0; i < filterNames.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(filterNames[i]))
                {
                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Warning,
                        $"Empty name found in Filter Names at position {i}.");
                    continue;
                }

                var filterName = filterNames[i].Trim();
                var newValue = newValues[i];

                if (nameToIndex.TryGetValue(filterName, out int index))
                {
                    // Variable found: apply change
                    resultAllValues[index] = newValue;
                    modifiedNames.Add(filterName);
                    modifiedValues.Add(newValue);

                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Remark,
                        $"Variable '{filterName}' changed: {allValues[index]:F3} → {newValue:F3}");
                }
                else
                {
                    // Variable not found
                    notFoundNames.Add(filterName);
                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Warning,
                        $"Variable '{filterName}' not found in All Names.");
                }
            }

            // Final statistics
            var totalModified = modifiedNames.Count;
            var totalNotFound = notFoundNames.Count;
            var totalRequested = filterNames.Count;

            if (totalModified > 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Remark,
                    $"✅ Modified {totalModified} of {totalRequested} variables. " +
                    $"Complete array ready for Play CPD ({resultAllValues.Count} values).");
            }

            if (totalNotFound > 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    $"⚠️ {totalNotFound} variables not found.");
            }

            // Set outputs
            DA.SetDataList(0, resultAllValues);    // → Connect directly to Play CPD
            DA.SetDataList(1, modifiedNames);      // For debug/visualization
            DA.SetDataList(2, modifiedValues);     // For debug/visualization  
            DA.SetDataList(3, notFoundNames);      // For debug/troubleshooting
        }

        public override Guid ComponentGuid
            => new Guid("A1F07C3D-4B8F-4E92-AB6C-DEADBEEF1234");

        protected override Bitmap Icon
            => Resources.Icon_SearchV;

        /// <summary>
        /// Additional component information
        /// </summary>
        public override string ToString()
        {
            return "GH_Calcpad_Search: Specific variable modifier";
        }
    }
}