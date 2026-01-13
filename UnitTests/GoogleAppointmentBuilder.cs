using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod.UnitTests
{
    public class GoogleAppointmentBuilder
    {
        public Event Build()
        {
            return Factory.NewEvent();
        }
    }
}
