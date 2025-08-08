using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component for filtering specific results from Calcpad calculations.
    /// Allows extracting only equations and values of interest from the complete result set.
    /// Workflow: Load → ModVar → Play → FilterResults → Export/Visualize
    /// </summary>
    public class GH_Calcpad_Search_Results : GH_Component
    {
        public GH_Calcpad_Search_Results()
          : base(
                "Search Results",       // Descriptive name
                "SearchRes",            // Short nickname
                "Filters specific equations and values from calculation results",
                "Calcpad",              // Category
                "5. Result Filtering"   // Filter Results
            )
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter(
                "Result Equations", "RE",
                "Complete list of resulting equations (from Play CPD)",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "Result Values", "RV",
                "Complete list of calculated values (from Play CPD)",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Filter Names", "FN",
                "Names of specific variables to extract from results",
                GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter(
                "Filtered Equations", "FE",
                "Filtered equations that match Filter Names",
                GH_ParamAccess.list);
            p.AddNumberParameter(
                "Filtered Values", "FV",
                "Values corresponding to filtered equations",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Found Names", "FN",
                "Variable names that were found successfully",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "Not Found", "NF",
                "Requested variables that were not found in results",
                GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var resultEquations = new List<string>();
            var resultValues = new List<double>();
            var filterNames = new List<string>();

            // Validate required inputs
            if (!DA.GetDataList(0, resultEquations)) return;
            if (!DA.GetDataList(1, resultValues)) return;
            if (!DA.GetDataList(2, filterNames)) return;

            // Validate input data consistency
            if (resultEquations.Count != resultValues.Count)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Error,
                    $"Result Equations ({resultEquations.Count}) and Result Values ({resultValues.Count}) must have the same length.");
                return;
            }

            // Validate there is data to process
            if (resultEquations.Count == 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    "No equations received in Result Equations.");
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, filterNames);
                return;
            }

            if (filterNames.Count == 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    "No variables specified for filtering.");
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<string>());
                return;
            }

            // Create dictionary for fast variable → index mapping
            var nameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < resultEquations.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(resultEquations[i]))
                {
                    string varName = ExtractVariableName(resultEquations[i]);
                    if (!string.IsNullOrEmpty(varName))
                    {
                        // If duplicates exist, keep the last index
                        nameToIndex[varName] = i;
                    }
                }
            }

            var filteredEquations = new List<string>();
            var filteredValues = new List<double>();
            var foundNames = new List<string>();
            var notFoundNames = new List<string>();

            // Process each requested variable
            foreach (var filterName in filterNames)
            {
                if (string.IsNullOrWhiteSpace(filterName))
                {
                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Warning,
                        "Empty name found in Filter Names.");
                    continue;
                }

                var cleanFilterName = filterName.Trim();

                if (nameToIndex.TryGetValue(cleanFilterName, out int index))
                {
                    // Variable found: add to filtered results
                    filteredEquations.Add(resultEquations[index]);
                    filteredValues.Add(resultValues[index]);
                    foundNames.Add(cleanFilterName);

                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Remark,
                        $"Variable '{cleanFilterName}' found: {resultValues[index]:F6}");
                }
                else
                {
                    // Variable not found
                    notFoundNames.Add(cleanFilterName);
                    AddRuntimeMessage(
                        GH_RuntimeMessageLevel.Warning,
                        $"Variable '{cleanFilterName}' not found in results.");
                }
            }

            // Final statistics
            var totalFound = foundNames.Count;
            var totalNotFound = notFoundNames.Count;
            var totalRequested = filterNames.Count;

            if (totalFound > 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Remark,
                    $"✅ Filtered {totalFound} of {totalRequested} requested variables.");
            }

            if (totalNotFound > 0)
            {
                AddRuntimeMessage(
                    GH_RuntimeMessageLevel.Warning,
                    $"⚠️ {totalNotFound} variables not found in results.");
            }

            // Set outputs
            DA.SetDataList(0, filteredEquations);   // Filtered equations
            DA.SetDataList(1, filteredValues);      // Filtered values
            DA.SetDataList(2, foundNames);          // Found variables
            DA.SetDataList(3, notFoundNames);       // Not found variables
        }

        /// <summary>
        /// Extracts variable name from the left side of an equation
        /// </summary>
        /// <param name="equation">Equation in format "variable = expression"</param>
        /// <returns>Variable name or empty string if cannot be extracted</returns>
        private string ExtractVariableName(string equation)
        {
            if (string.IsNullOrWhiteSpace(equation))
                return string.Empty;

            // Look for equals sign
            int equalIndex = equation.IndexOf('=');
            if (equalIndex <= 0)
                return string.Empty;

            // Extract left side and clean
            string leftSide = equation.Substring(0, equalIndex).Trim();

            // Validate it's a valid variable name (only letters, numbers, _, apostrophes)
            if (System.Text.RegularExpressions.Regex.IsMatch(leftSide, @"^[a-zA-Z_][a-zA-Z0-9_'′,\.]*$"))
            {
                return leftSide;
            }

            return string.Empty;
        }

        public override Guid ComponentGuid
            => new Guid("D4E5F6A7-B8C9-4D0E-1F2A-3B4C5D6E7F8A");

        protected override Bitmap Icon
            => Resources.Icon_SearchR;

        /// <summary>
        /// Additional component information
        /// </summary>
        public override string ToString()
        {
            return "GH_Calcpad_Filter_Results: Specific results filter";
        }

        /// <summary>
        /// Exposure level in interface
        /// </summary>
        public override GH_Exposure Exposure => GH_Exposure.primary;
    }
}