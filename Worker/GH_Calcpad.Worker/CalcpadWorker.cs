using System;
using System.Collections.Generic;
using Calcpad.Core;
using GH_Calcpad.Protocol;

namespace GH_Calcpad.Worker
{
    // Mirrors exactly what PyCalcpad.Parser.Parse(string code) does today
    // (new ExpressionParser{...}.Parse(code, true, false)) - no MacroParser,
    // no #include handling - since that's the only Calcpad entry point
    // GH_Calcpad's CalcpadSheet.Calculate() currently uses. Adding macro/include
    // preprocessing here would be new behavior, not a like-for-like replacement.
    internal static class CalcpadWorker
    {
        public static SolveResponse Solve(string id, string code)
        {
            var response = new SolveResponse { Id = id };
            try
            {
                var settings = new Settings();
                settings.Math.Decimals = 15;
                var parser = new ExpressionParser { Settings = settings };
                parser.Parse(code, true, false);

                // A variable reassigned more than once (e.g. recalculated inside a #for loop,
                // or refined further down the sheet) produces one ResultRecord per assignment,
                // in execution order. The LAST one is the variable's real final value - keep
                // that, not the first, while preserving first-seen order for a stable output.
                var byName = new Dictionary<string, ResultVariable>(StringComparer.Ordinal);
                var order = new List<string>();
                foreach (var record in parser.ResultRecords)
                {
                    if (!byName.ContainsKey(record.Name))
                        order.Add(record.Name);
                    byName[record.Name] = new ResultVariable
                    {
                        Name = record.Name,
                        Value = record.Value,
                        Unit = record.Unit
                    };
                }
                foreach (var name in order)
                    response.Results.Add(byName[name]);
                response.Ok = true;
            }
            catch (Exception ex)
            {
                response.Ok = false;
                response.Error = ex.Message;
            }
            return response;
        }
    }
}
