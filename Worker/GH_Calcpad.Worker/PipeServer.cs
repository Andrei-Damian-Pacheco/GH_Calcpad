using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using GH_Calcpad.Protocol;

namespace GH_Calcpad.Worker
{
    // One worker process per Rhino process (see migration plan): the pipe name
    // is passed in by the client (suffixed with the client's PID) so multiple
    // Rhino instances never collide. Accepts multiple concurrent connections;
    // each connection is NDJSON request/response, one SolveRequest per line in,
    // one SolveResponse per line out. Wire format lives in GH_Calcpad.Protocol
    // so client and server can never disagree on the JSON shape.
    //
    // Logs to stderr (the client redirects+captures it to a per-Rhino-process
    // log file) so a hang/crash leaves a trail: connection accepted, each
    // request id received, each response id sent.
    internal static class PipeServer
    {
        public static async Task RunAsync(string pipeName, CancellationToken token)
        {
            Log($"Listening on pipe '{pipeName}'.");
            while (!token.IsCancellationRequested)
            {
                var server = new NamedPipeServerStream(
                    pipeName,
                    PipeDirection.InOut,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                await server.WaitForConnectionAsync(token).ConfigureAwait(false);
                Log("Client connected.");

                // Handle this connection on its own task so the next client
                // (another concurrent session) isn't blocked waiting for it.
                _ = HandleConnectionAsync(server, token);
            }
        }

        private static async Task HandleConnectionAsync(NamedPipeServerStream server, CancellationToken token)
        {
            try
            {
                using (server)
                using (var reader = new StreamReader(server, System.Text.Encoding.UTF8))
                using (var writer = new StreamWriter(server, System.Text.Encoding.UTF8) { AutoFlush = false })
                {
                    while (server.IsConnected && !token.IsCancellationRequested)
                    {
                        var line = await reader.ReadLineAsync().ConfigureAwait(false);
                        if (line is null)
                            break;

                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var response = HandleRequestLine(line);
                        var json = Wire.Serialize(response);
                        await writer.WriteLineAsync(json).ConfigureAwait(false);
                        await writer.FlushAsync().ConfigureAwait(false);
                        Log($"Sent response id={response.Id} ok={response.Ok}");
                    }
                }
                Log("Client disconnected.");
            }
            catch (IOException ex)
            {
                Log($"Connection broken: {ex.Message}");
            }
            catch (ObjectDisposedException)
            {
            }
            catch (Exception ex)
            {
                // Never let a single connection's unexpected failure take down
                // the whole process silently.
                Log($"HandleConnectionAsync crashed: {ex}");
            }
        }

        private static SolveResponse HandleRequestLine(string line)
        {
            SolveRequest request;
            try
            {
                request = Wire.Deserialize<SolveRequest>(line);
            }
            catch (Exception ex)
            {
                Log($"Malformed request line: {ex.Message}");
                return new SolveResponse { Id = null, Ok = false, Error = $"Malformed request: {ex.Message}" };
            }

            if (request is null || string.IsNullOrEmpty(request.Code))
                return new SolveResponse { Id = request?.Id, Ok = false, Error = "Request must include 'code'." };

            bool isExport = string.Equals(request.Kind, "export", StringComparison.OrdinalIgnoreCase);
            Log($"Received {(isExport ? "export" : "solve")} request id={request.Id} ({request.Code.Length} chars).");
            try
            {
                return isExport
                    ? CalcpadExportWorker.Export(request.Id, request.Code, request.Format, request.OutputPath)
                    : CalcpadWorker.Solve(request.Id, request.Code);
            }
            catch (Exception ex)
            {
                // CalcpadWorker.Solve/CalcpadExportWorker.Export already catch Calcpad's own
                // exceptions; this is a last-resort net so a bug there can't hang the
                // connection - the caller always gets a response.
                Log($"Request id={request.Id} crashed: {ex}");
                return new SolveResponse { Id = request.Id, Ok = false, Error = $"Worker internal error: {ex.Message}" };
            }
        }

        private static void Log(string message) =>
            Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
    }
}
