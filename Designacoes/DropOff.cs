using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class DropOff
    {
        public DropOff(string name, string activity, string date, string email)
        {
            VolunteerName = name;
            ActivityName = activity;
            Date = date;
            Email = email;
        }

        public string ActivityName { get; set; }
        public string Email { get; set; }
        public string Date { get; set; }
        public string VolunteerName { get; set; }
    }
}
