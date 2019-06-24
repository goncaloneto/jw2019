using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Designacoes
{
    public class BusTrip
    {
        public string SlotName { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime ReturnTime { get; set; }
        public DateTime EndTime { get; set; }
        public int ActivityID { get; set; }
        public string Location { get; set; }
        public string ActivityName { get; set; }
        public int Delegates { get; set; }

        public string Observations { get; set; }

        public string Obs => String.IsNullOrEmpty(Observations) ? " " : Observations;

        public string StartTimeTime => StartTime.ToString("HH:mm").Replace(":", "h");
        public string ReturnTimeTime => ReturnTime.ToString("HH:mm").Replace(":", "h");
        public string StartTimeDate => StartTime.ToString("dd/MM/yyyy");

        public override bool Equals(Object obj)
        {
            //Check for null and compare run-time types.
            if ((obj == null) || !this.GetType().Equals(obj.GetType()))
            {
                return false;
            }
            else
            {
                BusTrip p = (BusTrip)obj;
                return (SlotName == p.SlotName) && (ActivityID == p.ActivityID) && (Location == p.Location) && (StartTime == p.StartTime);
            }
        }

        public class StartTimeComparer : IComparer<BusTrip>
        {
            public int Compare(BusTrip x, BusTrip y)
            {
                if (x == null || y == null)
                {
                    return 0;
                }

                // CompareTo() method 
                return x.StartTime.CompareTo(y.StartTime);

            }
        }
    }
}
