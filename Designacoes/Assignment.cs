using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class Assignment
    {
        public string SlotName { get; set; }
        public string VolunteerName { get; set; }
        public string VolunteerSurname { get; set; }
        public string Email { get; set; }
        public string Usage { get; set; }
        public DateTime Start { get; set; }
        public DateTime Return { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; }

        public string StartTime => Start.ToString("hh:mm").Replace(":","h");
        public string StartDate => Start.ToString("dd/MM/yyyy");
    }

    public class StartComparer : IComparer<Assignment>
    {
        public int Compare(Assignment x, Assignment y)
        {
            if (x == null || y == null)
            {
                return 0;
            }

            // CompareTo() method 
            return x.Start.CompareTo(y.Start);

        }
    }
}
