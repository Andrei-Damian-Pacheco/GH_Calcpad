using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component that provides guidance on the connection workflow between Grasshopper (GH) and Calcpad (CP).
    /// Explains typical workflows and best practices for using GH_Calcpad components.
    /// </summary>
    public class GH_Calcpad_Help : GH_Component
    {
        public GH_Calcpad_Help()
          : base("Calcpad Help", "CP_Help",
                 "Provides guidance on the connection workflow between Grasshopper and Calcpad",
                 "Calcpad", "7. Help & Support")
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            // No inputs required - this is an informational component
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Workflow", "W", "Step-by-step guide for typical workflows", GH_ParamAccess.list);
            pManager.AddTextParameter("ComponentGuide", "C", "Description of each component and its purpose", GH_ParamAccess.list);
            pManager.AddTextParameter("BestPractices", "B", "Best practices and usage tips", GH_ParamAccess.list);
            pManager.AddTextParameter("Examples", "E", "Common usage examples", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Updated workflows with new components
            var workflow = new List<string>
            {
                "MAIN WORKFLOWS:",
                "",
                "🔥 OPTIMIZATION WORKFLOW (REVOLUTIONARY):",
                "1. Load CPD → Optimizer → Galapagos → Best Solution → Save/Export",
                "   • Auto-detection of design variables",
                "   • Smart cache (5-10x faster)",
                "   • Zero configuration for typical cases",
                "",
                "📊 BASIC WORKFLOW:",
                "1. Load CPD/CPDz → Play → Export",
                "   • Load → Calculate → Report",
                "",
                "🔧 ADVANCED WORKFLOW:",
                "1. Load CPD → Modify Variables → Play → Filter Results → Export/Save",
                "   • Variables → Values → Calculate → Filter → Export",
                "",
                "📈 PARAMETRIC STUDY WORKFLOW:",
                "1. Load CPD → Series/Range → Play → Filter → Analysis",
                "   • Base → Parameter sets → Results → Analysis"
            };

            // 2) Updated guide with all components
            var componentGuide = new List<string>
            {
                "📋 BLOCK 1 - INFORMATION:",
                "• Info - Plugin and Calcpad.Core versions",
                "",
                "📁 BLOCK 2 - FILE LOADING:",
                "• Load CPD - .cpd files (source code)",
                "• Load CPDz - .cpdz files (compiled/protected)",
                "",
                "🔧 BLOCK 3 - MODIFICATION:",
                "• Modify Variables - Modifies variables pre-calculation",
                "",
                "⚡ BLOCK 4 - EXECUTION & OPTIMIZATION:",
                "• Play CPD - Main calculation engine",
                "• 🚀 Optimizer - AI for automatic optimization",
                "",
                "🔍 BLOCK 5 - FILTERING:",
                "• Filter Results - Extracts specific results",
                "",
                "💾 BLOCK 6 - SAVING & EXPORT:",
                "• Save CPD - Saves modified files",
                "• Export PDF - Vectorial PDF reports",
                "• Export Word - Editable Word documents",
                "",
                "❓ BLOCK 7 - HELP:",
                "• Help - This interactive guide"
            };

            // 3) Best practices with new features
            var bestPractices = new List<string>
            {
                "🚀 REVOLUTIONARY OPTIMIZATION:",
                "• Optimizer auto-detects design variables",
                "• Smart cache: 60x faster setup (30min → 30sec)",
                "• Native multi-objective with automatic convergence",
                "",
                "⚡ PERFORMANCE:",
                "• Cache system: 5-10x faster in optimization",
                "• CaptureExplicit=True for specific variables",
                "• Automatic file change monitoring",
                "",
                "🔧 VARIABLE HANDLING:",
                "• Maintain 1:1 correspondence between Variables/Values/Units",
                "• Use Modify Variables for selective changes",
                "• Filter Results to extract specific results",
                "",
                "📋 WORKFLOW ORGANIZATION:",
                "• Basic: Load → Play → Export",
                "• Advanced: Load → Modify → Play → Filter → Export",
                "• Optimization: Load → Optimizer → Galapagos → Save",
                "",
                "⚠️ ERROR HANDLING:",
                "• Check Success output in Play and Export",
                "• Optimizer includes convergence analysis",
                "• Automatic fallbacks for robustness"
            };

            // 4) Updated examples with new capabilities
            var examples = new List<string>
            {
                "🚀 AUTOMATIC OPTIMIZATION (REVOLUTIONARY):",
                "1. Load CPD → Optimizer (auto-detects variables)",
                "2. Connect to Galapagos (automatic configuration)",
                "3. Run optimization (smart cache)",
                "4. Save CPD + Export PDF (documentation)",
                "",
                "📊 ADVANCED PARAMETRIC ANALYSIS:",
                "1. Load CPD → Modify Variables",
                "2. Series → Play CPD → Filter Results",
                "3. Chart results + Export Word",
                "",
                "🔍 MULTI-OBJECTIVE OPTIMIZATION:",
                "1. Load CPD → Optimizer (specify objectives)",
                "2. Octopus/Galapagos multi-objetivo",
                "3. Pareto frontier analysis",
                "4. Best solution → Save + Export",
                "",
                "📈 COMPLETE AUTOMATED WORKFLOW:",
                "1. Load CPD → Modify Variables → Play",
                "2. Filter Results → Save CPD",
                "3. Export PDF + Export Word",
                "4. Timer → periodic automation",
                "",
                "⚙️ STRUCTURAL ENGINEERING:",
                "• Beam optimization: minimize weight + deflection",
                "• Column design: minimize cost + maximize safety",
                "• Truss optimization: minimize material + stress constraints"
            };

            // Set outputs
            DA.SetDataList(0, workflow);
            DA.SetDataList(1, componentGuide);
            DA.SetDataList(2, bestPractices);
            DA.SetDataList(3, examples);

            // Updated informational message
            AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, 
                $"GH_Calcpad Help v3.0 - {workflow.Count + componentGuide.Count + bestPractices.Count + examples.Count} lines | Includes revolutionary Optimizer");
        }

        public override Guid ComponentGuid
            => new Guid("A7B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D");

        protected override System.Drawing.Bitmap Icon
            => Resources.Icon_Help;

        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}
