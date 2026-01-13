using Google.Apis.Calendar.v3.Data;
using System.Collections.Generic;

namespace GoContactSyncMod
{
    internal class Factory
    {
        internal static Event NewEvent()
        {
            var ga = new Event
            {
                Reminders = new Event.RemindersData { Overrides = new List<EventReminder>(), UseDefault = false },
                Recurrence = new List<string>(),
                ExtendedProperties = new Event.ExtendedPropertiesData { Shared = new Dictionary<string, string>() },
                Start = new EventDateTime(),
                End = new EventDateTime(),
                Locked = false
            };
            return ga;
        }
    }
}
