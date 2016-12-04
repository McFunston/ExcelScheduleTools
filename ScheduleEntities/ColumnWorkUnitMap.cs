using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleEntities
{
    struct ColumnWorkUnitMap
    {
        public int JobNumber { get; set; }
        public int ClientName { get; set; }
        public int TaskName { get; set; }
        public int StartTime { get; set; }
        public int Duration { get; set; }
        public int Quantity { get; set; }
        public int IsReady { get; set; }
    }
}
