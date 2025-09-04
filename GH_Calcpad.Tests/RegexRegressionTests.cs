using Microsoft.VisualStudio.TestTools.UnitTesting;
using GH_Calcpad.Classes;

namespace GH_Calcpad.Tests
{
    [TestClass]
    public class RegexRegressionTests
    {
        [TestMethod]
        public void Alias_Tonf_To_kgf()
        {
            string s = "F = 1 tonf";
            string t = CalcpadSheet.NormalizeUnsupportedUnits(s);
            StringAssert.Contains(t, "1000 kgf");
        }

        [TestMethod]
        public void Alias_Kip_To_lbf()
        {
            string s = "F = 2 kip + 3 klbf";
            string t = CalcpadSheet.NormalizeUnsupportedUnits(s);
            StringAssert.Contains(t, "1000 lbf");
        }
    }
}