using System.Collections.Generic;

namespace GH_Calcpad.Protocol
{
    /// <summary>
    /// Wire request: full .cpd source text in, computed variables out (Kind="solve"),
    /// or a rendered file out (Kind="export", using Format/OutputPath). No session/load
    /// step exists (Phase 1 is output-side only - the caller is responsible for
    /// substituting input values into the source before sending it here), so every
    /// request is self-contained.
    /// </summary>
    public sealed class SolveRequest
    {
        public string Id { get; set; }
        public string Code { get; set; }

        /// <summary>"solve" (default, null/empty treated the same) or "export".</summary>
        public string Kind { get; set; }

        /// <summary>Export only: "html", "pdf" or "docx".</summary>
        public string Format { get; set; }

        /// <summary>Export only: full path of the file to write.</summary>
        public string OutputPath { get; set; }
    }

    public sealed class ResultVariable
    {
        public string Name { get; set; }
        public double Value { get; set; }
        public string Unit { get; set; }
    }

    public sealed class SolveResponse
    {
        public string Id { get; set; }
        public bool Ok { get; set; }
        public string Error { get; set; }
        public List<ResultVariable> Results { get; set; } = new List<ResultVariable>();

        /// <summary>Export only: the file that was actually written (== request's OutputPath on success).</summary>
        public string FinalPath { get; set; }
    }
}
