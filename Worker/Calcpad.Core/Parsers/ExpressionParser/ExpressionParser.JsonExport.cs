using System.Collections.Generic;

namespace Calcpad.Core
{
    public partial class ExpressionParser
    {
        // Added for GH_Calcpad's JSON worker (Phase 1, output side only).
        // Filled in ParseTokens (ExpressionParser.cs) as each visible assignment
        // line is calculated; cleared in Initialize() at the start of a fresh parse.
        public readonly record struct ResultRecord(string Name, double Value, string Unit);

        private readonly List<ResultRecord> _resultRecords = new();
        public IReadOnlyList<ResultRecord> ResultRecords => _resultRecords;
    }
}
