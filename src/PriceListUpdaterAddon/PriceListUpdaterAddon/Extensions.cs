using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PriceListUpdaterAddon
{
    internal static class Extensions
    {
        public static double ToDouble(this object o) => double.Parse(o.ToString());

        public static double NormalizeToThree(this double d) => Math.Round(d, 3);
    }
}
