using Grasshopper.Kernel;
using GH_Calcpad.Properties;

namespace GH_Calcpad
{
    /// <summary>
    /// Registers the "Calcpad" ribbon tab's own icon. GH_AssemblyInfo.Icon (see
    /// GH_CalcpadInfo.cs) only sets the icon shown in Grasshopper's plugin listing -
    /// the ribbon tab icon is a separate registration that Grasshopper never calls
    /// automatically, so without this the tab falls back to a default letter-in-circle
    /// icon derived from the category name ("C" for "Calcpad").
    /// </summary>
    public class GH_CalcpadPriority : GH_AssemblyPriority
    {
        public override GH_LoadingInstruction PriorityLoad()
        {
            Grasshopper.Instances.ComponentServer.AddCategoryIcon("Calcpad", Resources.Icon_Calcpad);
            Grasshopper.Instances.ComponentServer.AddCategorySymbolName("Calcpad", 'C');
            return GH_LoadingInstruction.Proceed;
        }
    }
}
