using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class Delegate
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string HotelName { get; set; }
        public string SlotName { get; set; }
        public string Language { get; set; }
    }

    public class HotelComparer : IComparer<Delegate>
    {
        public int Compare(Delegate x, Delegate y)
        {
            if (x == null || y == null)
            {
                return 0;
            }

            // CompareTo() method 
            return x.HotelName.CompareTo(y.HotelName);

        }
    }
}
