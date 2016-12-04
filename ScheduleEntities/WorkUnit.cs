using System;

namespace ScheduleEntities
{
    public class WorkUnit : IWorkUnit
    {
        public string JobNumber { get; set; }
        public string ClientName { get; set; }
        public string TaskName { get; set; }
        public DateTime StartTime { get; set; }
        public float Duration { get; set; }
        public int Quantity { get; set; }
        public bool IsReady { get; set; }
    }
}
