using System;

namespace ScheduleEntities
{
    interface IWorkUnit
    {
        string ClientName { get; set; }
        float Duration { get; set; }
        bool IsReady { get; set; }
        string JobNumber { get; set; }
        int Quantity { get; set; }
        DateTime StartTime { get; set; }
        string TaskName { get; set; }
    }
}