using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GH_Calcpad.Classes;
using GH_Calcpad.Properties;

namespace GH_Calcpad.Components
{
    /// <summary>
    /// Component that executes calculation on an existing CalcpadSheet
    /// applying new values and generating results.
    /// Simplified design: only Values + SheetObj as inputs.
    /// </summary>
    public class GH_Calcpad_Play : GH_Component
    {
        public GH_Calcpad_Play()
          : base("Play CPD", "PlayCPD",
                 "Executes calculation applying new values to SheetObj variables",
                 "Calcpad", "4. Execution & Optimization")  // Play CPD
        { }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddNumberParameter("Values", "V", "New values for variables (1:1 order with SheetObj Variables)", GH_ParamAccess.list);
            p.AddGenericParameter("SheetObj", "S", "CalcpadSheet from Load", GH_ParamAccess.item);

            p[0].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("ResultEq", "E", "Resulting equations", GH_ParamAccess.list);
            p.AddNumberParameter("ResultVal", "R", "Calculated numeric values", GH_ParamAccess.list);
            p.AddNumberParameter("Elapsed", "T", "Calculation time (ms)", GH_ParamAccess.item);
            p.AddBooleanParameter("Success", "S", "True if calculation successful", GH_ParamAccess.item);
            p.AddGenericParameter("UpdatedSheet", "US", "Updated CalcpadSheet for export", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            // 1) Read inputs
            var newValues = new List<double>();
            DA.GetDataList(0, newValues);

            object data = null;
            if (!DA.GetData(1, ref data))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No data received in SheetObj.");
                return;
            }

            // Unwrap CalcpadSheet
            CalcpadSheet sheet = null;
            if (data is GH_ObjectWrapper wrapper)
            {
                sheet = wrapper.Value as CalcpadSheet;
            }
            else
            {
                sheet = data as CalcpadSheet;
            }

            if (sheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"The received object is not a valid CalcpadSheet.");
                return;
            }

            AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                $"Debug - Variables: {sheet.Variables.Count}, {sheet.CodeInfo}");

            // 2) Validation: if Values provided, must match Variables
            if (newValues.Count > 0 && newValues.Count != sheet.Variables.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    $"Values ({newValues.Count}) must match SheetObj Variables ({sheet.Variables.Count}).");
                return;
            }

            try
            {
                // 3) Apply new values (if provided)
                if (newValues.Count > 0)
                {
                    for (int i = 0; i < sheet.Variables.Count; i++)
                    {
                        try
                        {
                            sheet.SetVariable(sheet.Variables[i], newValues[i]);
                        }
                        catch (Exception ex)
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                                              $"Could not assign '{sheet.Variables[i]}': {ex.Message}");
                        }
                    }
                }

                // 4) Execute calculation and measure time
                var sw = Stopwatch.StartNew();
                bool success = false;
                List<string> equations = new List<string>();
                List<double> results = new List<double>();

                try
                {
                    sheet.Calculate();
                    success = true;

                    equations = sheet.GetResultEquations();
                    results = sheet.GetResultValues();

                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                        $"Successful calculation: {equations.Count} equations");
                }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Calculation error: {ex.Message}");
                    equations = new List<string>();
                    results = new List<double>();
                }
                sw.Stop();

                // 5) Set outputs
                DA.SetDataList(0, equations);
                DA.SetDataList(1, results);
                DA.SetData(2, sw.Elapsed.TotalMilliseconds);
                DA.SetData(3, success);
                DA.SetData(4, new GH_ObjectWrapper(sheet));
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"General error: {ex.Message}");

                // Default outputs in case of error
                DA.SetDataList(0, new List<string>());
                DA.SetDataList(1, new List<double>());
                DA.SetData(2, 0.0);
                DA.SetData(3, false);
                DA.SetData(4, null);
            }
        }

        public override Guid ComponentGuid
            => new Guid("3B4A6ACA-3C2C-40E4-AB6C-ADACE17F78F5");

        protected override System.Drawing.Bitmap Icon
            => Resources.Icon_Play;
    }
}