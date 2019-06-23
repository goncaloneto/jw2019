using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class Slot
    {
        public string SlotName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Usage { get; set; }
        public string Location { get; set; }
        public DateTime Start { get; set; }

        public string StartDate => Start.ToString("dd/MM/yyyy");
    }
}
