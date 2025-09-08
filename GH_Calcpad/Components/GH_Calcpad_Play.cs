using System;
using System.Collections.Generic;
using System.Diagnostics;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    public class GH_Calcpad_Play : GH_Component
    {
        public GH_Calcpad_Play()
          : base("Play CPD", "PlayCPD",
                 "Executes calculation applying new values to SheetObj variables",
                 "Calcpad", "4. Execution & Optimization")
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddNumberParameter("Values", "V", "New values for variables (1:1 order with SheetObj Variables). Use NaN to skip a position.", GH_ParamAccess.list);
            p.AddGenericParameter("SheetObj", "S", "CalcpadSheet from Load", GH_ParamAccess.item);
            p[0].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("ResultEq", "E", "Result equations in 'Name=(...)' format", GH_ParamAccess.list);
            p.AddNumberParameter("ResultVal", "R", "Final numeric results", GH_ParamAccess.list);
            p.AddTextParameter("Units", "U", "Units of final results", GH_ParamAccess.list);
            p.AddNumberParameter("Elapsed", "T", "Calculation time (ms)", GH_ParamAccess.item);
            p.AddBooleanParameter("Success", "S", "True if calculation successful", GH_ParamAccess.item);
            p.AddGenericParameter("UpdatedSheet", "US", "Updated CalcpadSheet for export", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Inputs
            var newValues = new List<double>();
            DA.GetDataList(0, newValues);

            object data = null;
            if (!DA.GetData(1, ref data))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No data received in SheetObj.");
                return;
            }

            // Unwrap CalcpadSheet
            CalcpadSheet sheet = (data as GH_ObjectWrapper)?.Value as CalcpadSheet ?? data as CalcpadSheet;
            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The received object is not a valid CalcpadSheet.");
                return;
            }

            // 2) Apply values (tolerant)
            try
            {
                int varCount = sheet.Variables.Count;
                if (newValues.Count > 0)
                {
                    int min = Math.Min(varCount, newValues.Count);
                    if (newValues.Count != varCount)
                    {
                        if (newValues.Count > varCount)
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Received {newValues.Count} values, but Sheet has {varCount} variables. Extra values will be ignored.");
                        else
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Received {newValues.Count} values, but Sheet has {varCount} variables. Only first {min} variables will be updated.");
                    }

                    for (int i = 0; i < min; i++)
                    {
                        double v = newValues[i];
                        if (double.IsNaN(v) || double.IsInfinity(v)) continue;
                        TryAssign(sheet, sheet.Variables[i], v);
                    }
                }
            }
            catch (Exception exAssign)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error applying values: {exAssign.Message}");
            }

            // 3) Calculate
            var sw = Stopwatch.StartNew();
            bool success = false;
            var equations = new List<string>();
            var results = new List<double>();
            var units = new List<string>();

            try
            {
                sheet.Calculate();
                success = true;

                equations = sheet.GetResultEquations();
                results = sheet.GetResultValues();
                units = sheet.GetResultUnits();
            }
            catch (Exception exCalc)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Calculation error: {exCalc.Message}");
            }
            sw.Stop();

            // 4) Outputs
            DA.SetDataList(0, equations);
            DA.SetDataList(1, results);
            DA.SetDataList(2, units);
            DA.SetData(3, sw.Elapsed.TotalMilliseconds);
            DA.SetData(4, success);
            DA.SetData(5, new GH_ObjectWrapper(sheet));
        }

        private void TryAssign(CalcpadSheet sheet, string name, double value)
        {
            try { sheet.SetVariable(name, value); }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Could not assign '{name}': {ex.Message}");
            }
        }

        public override Guid ComponentGuid => new Guid("3B4A6ACA-3C2C-40E4-AB6C-ADACE17F78F5");
        protected override System.Drawing.Bitmap Icon => Resources.Icon_Play;
    }
}