using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GH_Calcpad.Protocol;

namespace GH_Calcpad.Classes
{
    /// <summary>
    /// Result of a single worker solve: every assignment line Calcpad computed
    /// while rendering the sheet, keyed by variable name. If a name is
    /// assigned more than once (e.g. recalculated inside a #for loop), the
    /// LAST occurrence's value/unit is kept - that's the variable's real
    /// final value.
    /// </summary>
    public sealed class CalcpadWorkerResult
    {
        public bool Ok;
        public string Error;
        public Dictionary<string, (double Value, string Unit)> Results
            = new Dictionary<string, (double, string)>(StringComparer.Ordinal);
    }

    /// <summary>
    /// Thin IPC client for the persistent GH_Calcpad.Worker process (Phase 1 of
    /// the JSON-worker migration - output side only). One worker process per
    /// Rhino process, spawned lazily on first use and torn down when this
    /// AppDomain unloads. See Worker\GH_Calcpad.Worker for the server side, and
    /// Worker\GH_Calcpad.Protocol for the shared wire contract used by both.
    /// </summary>
    public sealed class CalcpadWorkerClient : IDisposable
    {
        private static readonly Lazy<CalcpadWorkerClient> _instance =
            new Lazy<CalcpadWorkerClient>(() => new CalcpadWorkerClient());
        public static CalcpadWorkerClient Instance => _instance.Value;

        /// <summary>Where the worker's stdout/stderr get logged, for diagnosing hangs/crashes.</summary>
        public string LogFilePath { get; }

        private readonly object _connectLock = new object();
        private readonly ConcurrentDictionary<string, TaskCompletionSource<SolveResponse>> _pending =
            new ConcurrentDictionary<string, TaskCompletionSource<SolveResponse>>(StringComparer.Ordinal);

        private Process _process;
        private StreamWriter _logWriter;
        private NamedPipeClientStream _pipe;
        private StreamWriter _writer;
        private StreamReader _reader;
        private Task _readerLoop;
        private readonly object _writeLock = new object();

        private CalcpadWorkerClient()
        {
            var pid = Process.GetCurrentProcess().Id;
            LogFilePath = Path.Combine(Path.GetTempPath(), $"GH_Calcpad.Worker_{pid}.log");

            AppDomain.CurrentDomain.ProcessExit += (s, e) => Dispose();
            AppDomain.CurrentDomain.DomainUnload += (s, e) => Dispose();
        }

        /// <summary>
        /// Solves a full .cpd source text and returns every computed
        /// assignment's value+unit. Throws if the worker cannot be reached at
        /// all (caller is expected to surface a clear Grasshopper error, not
        /// fall back to any other calculation path).
        /// </summary>
        public CalcpadWorkerResult Solve(string code, int timeoutMs = 15000)
        {
            var request = new SolveRequest { Code = code };
            var response = SendRequestAsync(request, timeoutMs, retryOnce: true).GetAwaiter().GetResult();
            return ToWorkerResult(response);
        }

        /// <summary>
        /// Renders a full .cpd source text to HTML/PDF/DOCX exactly the way Calcpad's own
        /// CLI does (same worksheet template, same wkhtmltopdf.exe, same Calcpad.OpenXml
        /// writer - see CalcpadExportWorker on the worker side) and writes it to
        /// outputPath. Throws if the worker cannot be reached at all; returns Ok=false
        /// with Error set for any failure inside the export itself (bad format, wkhtmltopdf
        /// missing, etc.) so the caller can surface a clear message instead of an exception.
        /// </summary>
        public SolveResponse Export(string code, string format, string outputPath, int timeoutMs = 30000)
        {
            var request = new SolveRequest { Code = code, Kind = "export", Format = format, OutputPath = outputPath };
            return SendRequestAsync(request, timeoutMs, retryOnce: true).GetAwaiter().GetResult();
        }

        private async Task<SolveResponse> SendRequestAsync(SolveRequest request, int timeoutMs, bool retryOnce)
        {
            string id = null;
            try
            {
                // Connection failures (not just request timeouts) also get one
                // retry - a stale/dead pipe from a previous crashed worker should
                // not be a permanent failure.
                EnsureConnected();

                id = Guid.NewGuid().ToString("N");
                request.Id = id;
                var tcs = new TaskCompletionSource<SolveResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
                _pending[id] = tcs;

                var requestJson = Wire.Serialize(request);

                lock (_writeLock)
                {
                    _writer.WriteLine(requestJson);
                    _writer.Flush();
                }

                var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs)).ConfigureAwait(false);
                if (completed != tcs.Task)
                    throw new TimeoutException(
                        $"Calcpad worker did not respond within {timeoutMs} ms. Worker log: {LogFilePath}");

                return await tcs.Task.ConfigureAwait(false);
            }
            catch (Exception) when (retryOnce)
            {
                // Pipe likely broken (worker crashed) or never connected in the
                // first place. Reconnect once and retry this single request
                // before giving up.
                if (id != null) _pending.TryRemove(id, out _);
                Reset();
                request.Id = null;
                return await SendRequestAsync(request, timeoutMs, retryOnce: false).ConfigureAwait(false);
            }
            finally
            {
                if (id != null) _pending.TryRemove(id, out _);
            }
        }

        private static CalcpadWorkerResult ToWorkerResult(SolveResponse response)
        {
            var result = new CalcpadWorkerResult { Ok = response.Ok, Error = response.Error };
            foreach (var item in response.Results)
            {
                if (!string.IsNullOrEmpty(item.Name))
                    result.Results[item.Name] = (item.Value, item.Unit ?? string.Empty);
            }
            return result;
        }

        private void EnsureConnected()
        {
            if (_pipe != null && _pipe.IsConnected)
                return;

            lock (_connectLock)
            {
                if (_pipe != null && _pipe.IsConnected)
                    return;

                Reset();

                var pid = Process.GetCurrentProcess().Id;
                var pipeName = $"GH_Calcpad_Worker_{pid}";

                if (_process == null || _process.HasExited)
                    _process = StartWorkerProcess(pipeName, pid);

                var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                pipe.Connect(10000);

                _pipe = pipe;
                _writer = new StreamWriter(pipe, new UTF8Encoding(false)) { AutoFlush = false, NewLine = "\n" };
                _reader = new StreamReader(pipe, Encoding.UTF8);
                _readerLoop = Task.Run((Action)ReadLoop);
            }
        }

        private void ReadLoop()
        {
            try
            {
                string line;
                while ((line = _reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    SolveResponse response;
                    try
                    {
                        response = Wire.Deserialize<SolveResponse>(line);
                    }
                    catch
                    {
                        continue;
                    }

                    if (response?.Id != null && _pending.TryGetValue(response.Id, out var tcs))
                        tcs.TrySetResult(response);
                }
            }
            catch
            {
                // Connection broken; pending callers will time out and the
                // next Solve() call will reconnect via EnsureConnected/Reset.
            }
        }

        private Process StartWorkerProcess(string pipeName, int parentPid)
        {
            var exePath = ResolveWorkerExePath()
                ?? throw new FileNotFoundException("GH_Calcpad.Worker.exe not found next to the plugin.");

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"--pipe {pipeName} --parent-pid {parentPid}",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try { _logWriter?.Dispose(); } catch { }
            _logWriter = new StreamWriter(LogFilePath, append: false) { AutoFlush = true };
            _logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Starting {exePath} {psi.Arguments}");

            var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
            process.OutputDataReceived += (s, e) => { if (e.Data != null) LogSafe("[out] " + e.Data); };
            process.ErrorDataReceived += (s, e) => { if (e.Data != null) LogSafe("[err] " + e.Data); };
            process.Exited += (s, e) =>
            {
                LogSafe($"[client] worker process exited (code {SafeExitCode(process)}) while {_pending.Count} request(s) were pending.");
                // Fail any in-flight requests immediately instead of making
                // each one wait out its own timeout to discover the process
                // is gone.
                foreach (var kv in _pending)
                    kv.Value.TrySetException(new IOException(
                        $"GH_Calcpad.Worker.exe exited unexpectedly (code {SafeExitCode(process)}). See log: {LogFilePath}"));
            };

            if (!process.Start())
                throw new InvalidOperationException("Could not start GH_Calcpad.Worker.exe.");

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            return process;
        }

        private static int SafeExitCode(Process p)
        {
            try { return p.ExitCode; } catch { return -1; }
        }

        private void LogSafe(string line)
        {
            try { _logWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {line}"); } catch { }
        }

        private static string ResolveWorkerExePath()
        {
            var here = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(here))
                return null;

            // Lives in its own "Worker" subfolder, not flat next to the plugin -
            // see the BuildAndCopyWorker/PackageZip targets in GH_Calcpad.csproj.
            var local = Path.Combine(here, "Worker", "GH_Calcpad.Worker.exe");
            return File.Exists(local) ? local : null;
        }

        private void Reset()
        {
            try { _reader?.Dispose(); } catch { }
            try { _writer?.Dispose(); } catch { }
            try { _pipe?.Dispose(); } catch { }
            _reader = null;
            _writer = null;
            _pipe = null;

            foreach (var kv in _pending)
                kv.Value.TrySetException(new IOException("Calcpad worker connection was reset."));
            _pending.Clear();
        }

        public void Dispose()
        {
            Reset();
            try
            {
                if (_process != null && !_process.HasExited)
                    _process.Kill();
            }
            catch { }
            try { _logWriter?.Dispose(); } catch { }
        }
    }
}
