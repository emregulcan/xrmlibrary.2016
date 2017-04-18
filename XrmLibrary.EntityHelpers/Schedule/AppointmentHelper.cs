using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;
using XrmLibrary.EntityHelpers.Schedule.Data;

namespace XrmLibrary.EntityHelpers.Schedule
{
    /// <summary>
    /// An <c>Appointment</c> is a commitment that represents a time interval with start and end times and duration.
    /// The <c>Appointment</c> entity represents a block of time on a <c>Calendar</c>.
    /// This class provides mostly used common methods for <c>Appointment</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg327905(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class AppointmentHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Appointment.AddRecurrenceRequest
         * Unlike notices in SDK this message only copy "Subject", "Start date" and "End date" fields to new recurring appointment master instance.
         * In SDK notes you can see  "when you convert an existing appointment to a recurring appointment by using this message, the data from the existing appointment is copied to a new recurring appointment master instance", but this only copied "Subject" fields.
         * URL : https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addrecurrencerequest(v=crm.8).aspx
         */

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public AppointmentHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg327905(v=crm.8).aspx";
            this.EntityName = "appointment";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create an appointment from <see cref="XrmAppointment"/> object.
        /// </summary>
        /// <param name="appointment"></param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Create(XrmAppointment appointment)
        {
            ExceptionThrow.IfNull(appointment, "appointment");

            return this.OrganizationService.Create(appointment.ToEntity());
        }

        /// <summary>
        /// Add <c>recurrence</c> information to an existing <c>appointment</c>.
        /// Please note that when you convert an existing appointment to a recurring appointment, the data from the existing appointment is copied to a new recurring appointment master instance and the existing appointment record is deleted.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addrecurrencerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="appointmentId"><c>Appointment</c> Id</param>
        /// <param name="recurringAppointment">
        /// Recurring Appointment entity (<see cref="Entity"/>)
        /// </param>
        /// <returns>
        /// Newly created <c>Recurring Appointment</c> Id (<see cref="Guid"/>)
        /// </returns>
        public Guid AddRecurrence(Guid appointmentId, Entity recurringAppointment)
        {
            ExceptionThrow.IfGuidEmpty(appointmentId, "appointmentId");
            ExceptionThrow.IfNull(recurringAppointment, "recurringAppointment");
            ExceptionThrow.IfNullOrEmpty(recurringAppointment.LogicalName, "Entity.LogicalName");
            ExceptionThrow.IfNotExpectedValue(recurringAppointment.LogicalName, "Entity.LogicalName", "recurringappointmentmaster");

            AddRecurrenceRequest request = new AddRecurrenceRequest()
            {
                AppointmentId = appointmentId,
                Target = recurringAppointment
            };

            return ((AddRecurrenceResponse)this.OrganizationService.Execute(request)).id;
        }

        /// <summary>
        /// Add <c>recurrence</c> information to an existing <c>appointment</c> with <see cref="XrmRecurringAppointment"/> object.
        /// Please note that when you convert an existing appointment to a recurring appointment, the data from the existing appointment is copied to a new recurring appointment master instance and the existing appointment record is deleted.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addrecurrencerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="appointmentId"><c>Appointment</c> Id</param>
        /// <param name="recurringAppointment">
        /// <see cref="XrmRecurringAppointment"/>
        /// </param>
        /// <returns>
        /// Newly created <c>Recurring Appointment</c> Id (<see cref="Guid"/>)
        /// </returns>
        public Guid AddRecurrence(Guid appointmentId, XrmRecurringAppointment recurringAppointment)
        {
            ExceptionThrow.IfGuidEmpty(appointmentId, "appointmentId");

            AddRecurrenceRequest request = new AddRecurrenceRequest()
            {
                AppointmentId = appointmentId,
                Target = recurringAppointment.ToEntity()
            };

            return ((AddRecurrenceResponse)this.OrganizationService.Execute(request)).id;
        }

        /// <summary>
        /// Schedule or “book” an appointment.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.bookrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="appointment"></param>
        /// <returns>
        /// Returns <c>Validation Result</c> (<see cref="ValidationResult"/>) in <see cref="ValidateResponse"/>.
        /// You can reach created <c>Appointment</c> Id in <see cref="ValidationResult.ActivityId"/> property.
        /// </returns>
        public BookResponse Book(XrmAppointment appointment)
        {
            ExceptionThrow.IfNull(appointment, "appointment");

            BookRequest request = new BookRequest()
            {
                Target = appointment.ToEntity()
            };

            return (BookResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// <c>Reschedule</c> an appointment.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.reschedulerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="appointmentId"></param>
        /// <param name="start">Start time</param>
        /// <param name="end">End time</param>
        /// <returns>
        /// <see cref="RescheduleResponse"/>
        /// </returns>
        public RescheduleResponse Reschedule(Guid appointmentId, DateTime start, DateTime end)
        {
            ExceptionThrow.IfGuidEmpty(appointmentId, "appointmentId");
            ExceptionThrow.IfEquals(start, "start", DateTime.MinValue);
            ExceptionThrow.IfEquals(start, "start", DateTime.MaxValue);
            ExceptionThrow.IfEquals(end, "end", DateTime.MinValue);
            ExceptionThrow.IfEquals(end, "end", DateTime.MaxValue);

            Entity target = new Entity("appointment");
            target["activityid"] = appointmentId;
            target["scheduledstart"] = start;
            target["scheduledend"] = end;

            RescheduleRequest request = new RescheduleRequest()
            {
                Target = target
            };

            return (RescheduleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Verify that an appointment has valid available resources for the activity, duration, and site, as appropriate.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.validaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="idList">Activity Id list</param>
        /// <returns>
        /// <see cref="ValidateResponse"/>
        /// Returns <c>Validation Result</c> (<see cref="ValidationResult"/>) in <see cref="ValidateResponse.Result"/>.
        /// </returns>
        public ValidateResponse Validate(List<Guid> idList)
        {
            var activities = new EntityCollection();

            foreach (var item in idList)
            {
                var entity = this.Get(item, new string[] { "scheduledstart", "scheduledend", "statecode", "statuscode" });
                activities.Entities.Add(entity);
            }

            ValidateRequest request = new ValidateRequest()
            {
                Activities = activities
            };

            return (ValidateResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
