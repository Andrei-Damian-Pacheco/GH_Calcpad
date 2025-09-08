using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Help component for GH_Calcpad.
    /// Provides workflows, component overview, best practices and usage examples.
    /// Updated for version 1.2.0 (current supported input: .cpd; .cpdz experimental / limited).
    /// All text in English for documentation consistency.
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

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Workflows (only .cpd fully supported in v1.2.0)
            var workflow = new List<string>
            {
                "MAIN WORKFLOWS (v1.2.0 – current support: .cpd):",
                "",
                "🔥 OPTIMIZATION WORKFLOW:",
                "1. Load CPD → Optimizer → Galapagos → Best Solution → Save / Export",
                "   • Automatic detection of design variables",
                "   • Single fitness + objective values",
                "",
                "📊 BASIC WORKFLOW:",
                "1. Load CPD → Play → Export",
                "   • Load → Compute → Report",
                "",
                "🔧 VARIABLE MODIFICATION WORKFLOW:",
                "1. Load CPD → Search Variables → Play → Export",
                "   • Selectively modify chosen variables, then compute",
                "",
                "📈 ADVANCED (RESULT FILTERING):",
                "1. Load CPD → (Search Variables optional) → Play → Search Results → Export / Save",
                "   • Modify subset → Compute → Filter key outputs",
                "",
                "🧪 PARAMETRIC STUDY:",
                "1. Load CPD → (Sliders / Series / Range) → Search Variables → Play → Search Results → Analyze",
                "   • Generate variants → Evaluate → Visualize",
                "",
                "ℹ NOTE: Load CPDz component exists; full compiled (.cpdz) workflow support is planned for a future release."
            };

            // 2) Component guide
            var componentGuide = new List<string>
            {
                "📋 INFORMATION & DIAGNOSTICS:",
                "• Info – Plugin + Calcpad.Core versions",
                "",
                "📁 FILE LOADING:",
                "• Load CPD – Load and parse .cpd source sheets",
                "• Load CPDz – Experimental (works only if textual source is embedded)",
                "",
                "🔧 VARIABLE MODIFICATION:",
                "• Search Variables – Filter + overwrite selected variables keeping full order",
                "",
                "⚡ EXECUTION & OPTIMIZATION:",
                "• Play CPD – Core calculation engine",
                "• Optimizer – Multi-objective fitness preparation + caching",
                "",
                "🔍 RESULT FILTERING:",
                "• Search Results – Extract specific result variables",
                "",
                "💾 SAVING & EXPORT:",
                "• Save CPD – Save modified sheet (.cpd / .txt)",
                "• Export HTML – HTML report",
                "• Export PDF – PDF report",
                "• Export Word – Editable .docx report",
                "",
                "❓ HELP & SUPPORT:",
                "• Help – This guide"
            };

            // 3) Best practices
            var bestPractices = new List<string>
            {
                "⚙ VARIABLES:",
                "• Maintain 1:1 alignment: Variables / Values / Units.",
                "• Use Search Variables for partial updates (do not reorder original lists).",
                "• Use NaN in 'Values' (Play) to skip a variable position.",
                "",
                "🚀 OPTIMIZER:",
                "• Leave design variable list empty first run → auto-detection.",
                "• Review 'Status' + 'Convergence Info' to decide stopping.",
                "• Reuse the same sheet instance to leverage caching.",
                "",
                "📐 PERFORMANCE:",
                "• Freeze unchanged upstream elements to avoid recompute.",
                "• Use CaptureExplicit=True when you only need explicit marked variables.",
                "",
                "🔍 RESULTS:",
                "• Search Results to reduce downstream graph clutter.",
                "• Keep result variable names short and without spaces.",
                "",
                "📦 WORKFLOWS:",
                "• Basic: Load → Play → Export.",
                "• Advanced: Load → Search Variables → Play → Search Results → Export.",
                "• Optimization: Load → Optimizer → Galapagos → Save / Export.",
                "",
                "🛠 ERROR HANDLING:",
                "• Always check 'Success' outputs.",
                "• Yellow = warning (recoverable), Red = fix required.",
                "• If a value not updated: verify exact variable name.",
                "",
                "📄 EXPORT:",
                "• Reuse 'UpdatedSheet' across HTML/PDF/Word components.",
                "• Save CPD before exports if internal state changed.",
                "",
                "🔮 FUTURE:",
                "• Enhanced .cpdz support + richer optimizer diagnostics planned."
            };

            // 4) Examples
            var examples = new List<string>
            {
                "🚀 AUTOMATIC OPTIMIZATION:",
                "1. Load CPD → Optimizer → Galapagos → Best fitness → Save CPD → Export PDF",
                "",
                "🎯 MULTI-OBJECTIVE:",
                "1. Load CPD → Optimizer (set objectives/modes) → Galapagos / Octopus → Export Word + PDF",
                "",
                "🔧 SELECTIVE EDIT:",
                "1. Load CPD → Search Variables (change few params) → Play → Export HTML",
                "",
                "📊 PARAMETRIC STUDY:",
                "1. Load CPD → Sliders / Series → Search Variables → Play → Search Results → Graph",
                "",
                "🗂 COMPLETE PIPELINE:",
                "1. Load CPD → Search Variables → Play → Search Results → Save CPD → Export PDF / Word",
                "",
                "🏗 STRUCTURAL:",
                "• Minimize weight with displacement + stress constraints",
                "• Column optimization: cost vs slenderness",
                "",
                "🛠 MECHANICAL:",
                "• Section tuning: stiffness vs mass",
                "• Thermal/material parameter sweeps",
                "",
                "🏛 ENVELOPE:",
                "• Insulation parameter variations → energy performance comparison",
                "• Compare alternatives using filtered key outputs"
            };

            // Set outputs
            DA.SetDataList(0, workflow);
            DA.SetDataList(1, componentGuide);
            DA.SetDataList(2, bestPractices);
            DA.SetDataList(3, examples);

            AddRuntimeMessage(
                GH_RuntimeMessageLevel.Remark,
                $"GH_Calcpad Help v1.2.0 | Lines: {workflow.Count + componentGuide.Count + bestPractices.Count + examples.Count}");
        }

        public override Guid ComponentGuid
            => new Guid("A7B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D");

        protected override System.Drawing.Bitmap Icon
            => Resources.Icon_Help;

        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}
