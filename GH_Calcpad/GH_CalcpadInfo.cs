using GH_Calcpad.Properties;
using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;
using GH_Calcpad.Classes;

namespace GH_Calcpad
{
    public class GH_CalcpadInfo : GH_AssemblyInfo
    {
        static GH_CalcpadInfo()
        {
            // Hook del resolver al cargar el ensamblado
            AssemblyResolver.EnsureHook();
        }

        public override string Name => "GH_Calcpad";

        /// <summary>
        /// Return a 24x24 pixel bitmap to represent this GHA library in Grasshopper tabs.
        /// </summary>
        public override Bitmap Icon
        {
            get
            {
                try
                {
                    // Usar el ICO principal del proyecto y convertir a Bitmap
                    return Resources.GH_Calcpad?.ToBitmap();
                }
                catch
                {
                    // Fallback: usar el PNG si falla el ICO
                    return Resources.Icon_Calcpad;
                }
            }
        }

        public override string Description => "Integración completa de Calcpad en Grasshopper con optimización avanzada";
        public override Guid Id => new Guid("1A9A80A9-0512-4001-8A73-1FECB82117B7");
        public override string AuthorName => "Andrei Damian Pacheco";
        public override string AuthorContact => "ad.andreidamian@gmail.com";
        public override string AssemblyVersion => GetType().Assembly.GetName().Version.ToString();
    }
}