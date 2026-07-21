using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using GH_Calcpad.Worker;

// Two modes:
//   GH_Calcpad.Worker <path-to-cpd-file>
//     One-shot smoke test: solve a file and print the JSON. Used during
//     development to check the JSON dump against the real Calcpad app.
//   GH_Calcpad.Worker --pipe <name> --parent-pid <pid>
//     Persistent named-pipe server, launched by GH_Calcpad. Exits on its own
//     if the parent (Rhino) process disappears, so it never survives an
//     ungraceful Rhino shutdown.

if (args.Length >= 2 && args[0] == "--pipe")
{
    string pipeName = args[1];
    int parentPid = -1;
    for (int i = 2; i < args.Length - 1; i++)
        if (args[i] == "--parent-pid")
            int.TryParse(args[i + 1], out parentPid);

    using var cts = new CancellationTokenSource();

    if (parentPid > 0)
        _ = MonitorParentAsync(parentPid, cts);

    try
    {
        await PipeServer.RunAsync(pipeName, cts.Token);
    }
    catch (OperationCanceledException)
    {
    }
    catch (Exception ex)
    {
        // Last-resort log so a crash in the accept loop itself (outside any
        // single connection) still leaves a trail in the client's log file
        // instead of the process just silently vanishing.
        Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] FATAL: PipeServer.RunAsync crashed: {ex}");
        return 3;
    }
    return 0;
}

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: GH_Calcpad.Worker <path-to-cpd-file>");
    Console.Error.WriteLine("       GH_Calcpad.Worker --pipe <name> --parent-pid <pid>");
    return 1;
}

var code = File.ReadAllText(args[0]);
var result = CalcpadWorker.Solve(Guid.NewGuid().ToString(), code);
var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
{
    WriteIndented = true,
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
});
Console.WriteLine(json);
return result.Ok ? 0 : 2;

static async Task MonitorParentAsync(int parentPid, CancellationTokenSource cts)
{
    while (!cts.IsCancellationRequested)
    {
        try
        {
            Process.GetProcessById(parentPid);
        }
        catch (ArgumentException)
        {
            cts.Cancel();
            return;
        }
        await Task.Delay(2000, cts.Token).ConfigureAwait(false);
    }
}
