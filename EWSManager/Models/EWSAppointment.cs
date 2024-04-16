using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;


namespace EWSManager.Models
{
    public class EWSAppointment
    {
        public string subject { get; set; }
        public string body { get; set; }
        public DateTime startDate { get; set; }
        public int duration { get; set; }
        public string? location { get; set; }
        public List<string>? attendees { get; set; }


        public void Create(EWSManager eWSManager) 
        {
            Appointment appointment = new Appointment(eWSManager.exchangeService);

            appointment.Subject = subject;
            appointment.Body = body;
            appointment.Start = startDate;
            appointment.End = appointment.Start.AddHours(duration);
            if (location != null )
            {
                appointment.Location = location;

            }
            if (attendees != null )
            {
                foreach ( var a in attendees )
                    appointment.RequiredAttendees.Add(a);
            }

            appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
        }
    }
}
