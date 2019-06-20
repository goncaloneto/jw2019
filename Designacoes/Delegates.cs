using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class Delegate
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Hotel { get; set; }
        public string SlotName { get; set; }
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
            return x.Hotel.CompareTo(y.Hotel);

        }
    }
}
