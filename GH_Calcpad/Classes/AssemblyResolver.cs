using System;
using System.IO;
using System.Reflection;

namespace GH_Calcpad.Classes
{
    internal static class AssemblyResolver
    {
        private static bool _hooked;

        internal static void EnsureHook()
        {
            if (_hooked) return;
            AppDomain.CurrentDomain.AssemblyResolve += Resolve;
            _hooked = true;
        }

        private static Assembly Resolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var name = new AssemblyName(args.Name).Name + ".dll";

                // 1) Plugin folder (the .gha directory)
                var pluginDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var local = Path.Combine(pluginDir ?? string.Empty, name);
                if (File.Exists(local))
                    return Assembly.LoadFrom(local);

                // 2) Calcpad installation (default under Program Files)
                var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                var calcpadDir = Path.Combine(programFiles, "Calcpad");
                var fromCalcpad = Path.Combine(calcpadDir, name);
                if (File.Exists(fromCalcpad))
                    return Assembly.LoadFrom(fromCalcpad);

                // 3) Program Files (x86) fallback (in case of 32‑bit install)
                var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                var calcpadDirX86 = Path.Combine(programFilesX86, "Calcpad");
                var fromCalcpadX86 = Path.Combine(calcpadDirX86, name);
                if (File.Exists(fromCalcpadX86))
                    return Assembly.LoadFrom(fromCalcpadX86);
            }
            catch
            {
                // Silent: let normal resolution continue
            }

            return null;
        }
    }
}