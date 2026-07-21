using System.Text.Json;

namespace GH_Calcpad.Protocol
{
    /// <summary>
    /// Single source of truth for the NDJSON wire format, shared by the net48
    /// client (GH_Calcpad.Classes.CalcpadWorkerClient) and the net10 worker
    /// (GH_Calcpad.Worker), so the two sides can never silently disagree on
    /// the JSON shape.
    /// </summary>
    public static class Wire
    {
        public static readonly JsonSerializerOptions Options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        public static string Serialize<T>(T value) => JsonSerializer.Serialize(value, Options);

        public static T Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, Options);
    }
}
