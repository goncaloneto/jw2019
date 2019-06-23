using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class DropOff
    {
        public DropOff(string name, string activity, string date)
        {
            VolunteerName = name;
            ActivityName = activity;
            Date = date;
        }

        public string ActivityName { get; set; }
        public string Date { get; set; }
        public string VolunteerName { get; set; }
    }
}
