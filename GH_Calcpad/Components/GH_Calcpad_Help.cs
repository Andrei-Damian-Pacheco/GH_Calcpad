using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Help component for GH_Calcpad.
    /// Provides workflows, component overview, best practices and usage examples.
    /// Updated for version 2.0.0 (worker-based architecture).
    /// </summary>
    public class GH_Calcpad_Help : GH_Component
    {
        public GH_Calcpad_Help()
          : base("Calcpad Help", "CP_Help",
                 "Provides workflows, component overview, best practices and examples for GH_Calcpad",
                 "Calcpad", "7. Help & Support")
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            // No inputs
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Workflow", "W", "Step-by-step guides for the main workflows", GH_ParamAccess.list);
            pManager.AddTextParameter("ComponentGuide", "C", "Description of each component and purpose", GH_ParamAccess.list);
            pManager.AddTextParameter("BestPractices", "B", "Best practices and recommendations", GH_ParamAccess.list);
            pManager.AddTextParameter("Examples", "E", "Common usage examples", GH_ParamAccess.list);
        }

        // Contenido 100% estático (sin inputs): se calcula una sola vez por sesión de Rhino y se
        // comparte entre instancias, en vez de reconstruir las 4 listas en cada SolveInstance.
        private static readonly Lazy<List<string>> _workflow = new Lazy<List<string>>(BuildWorkflow);
        private static readonly Lazy<List<string>> _componentGuide = new Lazy<List<string>>(BuildComponentGuide);
        private static readonly Lazy<List<string>> _bestPractices = new Lazy<List<string>>(BuildBestPractices);
        private static readonly Lazy<List<string>> _examples = new Lazy<List<string>>(BuildExamples);

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var workflow = _workflow.Value;
            var componentGuide = _componentGuide.Value;
            var bestPractices = _bestPractices.Value;
            var examples = _examples.Value;

            DA.SetDataList(0, workflow);
            DA.SetDataList(1, componentGuide);
            DA.SetDataList(2, bestPractices);
            DA.SetDataList(3, examples);

            AddRuntimeMessage(
                GH_RuntimeMessageLevel.Remark,
                $"GH_Calcpad Help v2.0.0 | Lines: {workflow.Count + componentGuide.Count + bestPractices.Count + examples.Count}");
        }

        private static List<string> BuildWorkflow() => new List<string>
        {
            "Basic:        Load CPD -> Play CPD -> Export/Save",
            "Selective:    Load CPD -> Search Variables -> Play CPD -> Export/Save",
            "Full:         Load CPD -> Search Variables -> Play CPD -> Search Results -> Export/Save",
            "Optimization: same as Full, leave both 'Filter Names' empty (auto gh_/ghc_), wire into Galapagos/Wallacei/Octopus"
        };

        private static List<string> BuildComponentGuide() => new List<string>
        {
            "Load CPD: FilePath, Freeze -> Variables, Values, Units, Sheet",
            "Search Variables: Sheet, Filter Names (empty = auto 'gh_'), New Values -> Sheet, Modified Names/Values, Not Found",
            "Play CPD: Sheet -> Sheet (calculated), Elapsed, Success",
            "Search Results: Sheet, Filter Names (empty = auto 'ghc_') -> Equations, Values, Units",
            "Save CPD/TXT: writes Sheet's code to .cpd/.txt",
            "Export HTML/PDF/Word: renders Sheet exactly like Calcpad's own app/CLI",
            "Calcpad Info: plugin + Calcpad.Core versions",
            "Calcpad Help: this guide"
        };

        private static List<string> BuildBestPractices() => new List<string>
        {
            "Prefix design variables 'gh_' and result variables 'ghc_' to use auto-detect in Search Variables/Search Results.",
            "Use the real prime character (′), never a straight apostrophe ('): in Calcpad, ' is always a comment.",
            "Check 'Success' on Play CPD and on Export/Save before trusting downstream results.",
            "'Not Found' on Search Variables/Search Results flags a typo'd or missing variable name."
        };

        private static List<string> BuildExamples() => new List<string>
        {
            "Optimization: Load CPD -> Search Variables (Filter Names empty, sliders into New Values) -> Play CPD -> Search Results (Filter Names empty) -> Galapagos fitness",
            "Report only:  Load CPD -> Play CPD -> Export PDF"
        };

        public override Guid ComponentGuid
            => new Guid("A7B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D");

        protected override System.Drawing.Bitmap Icon
            => Resources.Icon_Help;

        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}
