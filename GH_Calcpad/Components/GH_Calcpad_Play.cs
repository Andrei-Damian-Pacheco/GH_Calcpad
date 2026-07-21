using System;
using System.Diagnostics;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Calculation engine. Takes a fully-prepared Sheet (values already set via Search
    /// Variables, or plain literals straight from Load CPD) and runs it through Calcpad.
    /// The returned Sheet feeds both Search Results (reads the computed values) and any
    /// Export component (reads OriginalCode, untouched by Calculate()).
    /// </summary>
    public class GH_Calcpad_Play : GH_Component
    {
        public GH_Calcpad_Play()
          : base("Play CPD", "PlayCPD",
                 "Calculates a CalcpadSheet",
                 "Calcpad", "4. Execution & Optimization")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "CalcpadSheet to calculate (from Load CPD or Search Variables)", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddGenericParameter("Sheet", "S", "Calculated CalcpadSheet - connect to Search Results and/or Export", GH_ParamAccess.item);
            p.AddNumberParameter("Elapsed", "T", "Calculation time (ms)", GH_ParamAccess.item);
            p.AddBooleanParameter("Success", "OK", "True if calculation succeeded", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object data = null;
            if (!DA.GetData(0, ref data))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No data received in Sheet.");
                return;
            }

            CalcpadSheet incomingSheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (incomingSheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return;
            }

            // Clone before mutating: Calculate() writes into the Sheet's internal result
            // cache, and the same upstream Sheet could be wired into more than one Play CPD.
            CalcpadSheet sheet = incomingSheet.Clone();

            var sw = Stopwatch.StartNew();
            bool success = false;
            try
            {
                sheet.Calculate();
                success = true;
            }
            catch (Exception exCalc)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Calculation error: {exCalc.Message}");
            }
            sw.Stop();

            DA.SetData(0, new GH_ObjectWrapper(sheet));
            DA.SetData(1, sw.Elapsed.TotalMilliseconds);
            DA.SetData(2, success);
        }

        public override Guid ComponentGuid => new Guid("3B4A6ACA-3C2C-40E4-AB6C-ADACE17F78F5");
        protected override System.Drawing.Bitmap Icon => Resources.Icon_Play;
    }
}
