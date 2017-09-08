using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;

namespace XrmLibrary.EntityHelpers.Schedule.Data
{
    /// <summary>
    /// Fluent object to easy-create MS CRM <c>Appointment</c> activity.
    /// </summary>
    public class XrmAppointment
    {
        #region | Private Definitions |

        bool _isInit = false;
        int _priority;

        KeyValuePair<AttendeeEntityTypeCode, Guid> _organizer;
        KeyValuePair<string, Guid> _regardingObject;
        List<Tuple<AttendeeTypeCode, AttendeeEntityTypeCode, Guid>> _attendeeList;
        Dictionary<String, object> _customAttributeList;

        string _subject;
        string _description;
        string _location;
        DateTime? _start;
        DateTime? _end;
        bool _isAllDayEvent;

        #endregion

        #region | Enums |

        /// <summary>
        /// StateCode : 0
        /// </summary>
        public enum AppointmentActiveStatusCode
        {
            Free = 1,
            Tentative = 2
        }

        /// <summary>
        /// StateCode : 1
        /// </summary>
        public enum AppointmentCompleteStatusCode
        {
            Completed = 3
        }

        /// <summary>
        /// StateCode : 2
        /// </summary>
        public enum AppointmentCancelStatusCode
        {
            Cancelled = 4
        }

        /// <summary>
        /// StateCode : 3
        /// </summary>
        public enum AppointmentScheduledStatusCode
        {
            Busy = 5,
            OutOfOffice = 6
        }

        public enum AttendeeTypeCode
        {
            Required = 1,
            Optional = 2
        }

        public enum AttendeeEntityTypeCode
        {
            [Description("lead")]
            Lead,

            [Description("account")]
            Account,

            [Description("contact")]
            Contact,

            [Description("equipment")]
            Equipment,

            [Description("systemuser")]
            SystemUser
        }

        public enum PriorityCode
        {
            Low = 0,
            Normal = 1,
            High = 2
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// 
        /// </summary>
        public XrmAppointment()
        {
            _isInit = true;
            _priority = (int)PriorityCode.Normal;
            _attendeeList = new List<Tuple<AttendeeTypeCode, AttendeeEntityTypeCode, Guid>>();
            _customAttributeList = new Dictionary<string, object>();
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add <c>attendee</c>
        /// </summary>
        /// <param name="typecode"><see cref="AttendeeTypeCode"/></param>
        /// <param name="entityTypeCode"><see cref="AttendeeEntityTypeCode"/></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment AddAttendee(AttendeeTypeCode typecode, AttendeeEntityTypeCode entityTypeCode, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            _attendeeList.Add(new Tuple<AttendeeTypeCode, AttendeeEntityTypeCode, Guid>(typecode, entityTypeCode, id));

            return this;
        }

        /// <summary>
        /// Add <c>organizer</c>
        /// </summary>
        /// <param name="entityTypeCode"><see cref="AttendeeEntityTypeCode"/></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment AddOrganizer(AttendeeEntityTypeCode entityTypeCode, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            _organizer = new KeyValuePair<AttendeeEntityTypeCode, Guid>(entityTypeCode, id);
            return this;
        }

        /// <summary>
        /// Add <c>subject</c>
        /// </summary>
        /// <param name="subject">Subject</param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment Subject(string subject)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            _subject = subject.Trim();
            return this;
        }

        /// <summary>
        /// Add <c>description</c>
        /// </summary>
        /// <param name="description">Description</param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment Description(string description)
        {
            ExceptionThrow.IfNullOrEmpty(description, "description");

            _description = description.Trim();
            return this;
        }

        /// <summary>
        /// Add <c>location</c>
        /// </summary>
        /// <param name="location">Location</param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment Location(string location)
        {
            ExceptionThrow.IfNullOrEmpty(location, "location");

            _location = location.Trim();
            return this;
        }

        /// <summary>
        /// Is appointment <c>all day event</c>?
        /// </summary>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment AllDayEvent()
        {
            _isAllDayEvent = true;
            return this;
        }

        /// <summary>
        /// Appointment <c>Start date</c>
        /// </summary>
        /// <param name="date"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment ScheduledStart(DateTime date)
        {
            ExceptionThrow.IfEquals(date, "ScheduledStart", DateTime.MinValue);

            _start = date;
            return this;
        }

        /// <summary>
        /// Appointment <c>End date</c>
        /// </summary>
        /// <param name="date"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment ScheduledEnd(DateTime date)
        {
            ExceptionThrow.IfEquals(date, "ScheduledEnd", DateTime.MinValue);

            _end = date;
            return this;
        }

        /// <summary>
        /// Add <c>priority</c>
        /// </summary>
        /// <param name="priority"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment Priority(PriorityCode priority)
        {
            _priority = (int)priority;
            return this;
        }

        /// <summary>
        /// Add <c>regarding</c>
        /// </summary>
        /// <param name="entityLogicalName"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmAppointment"/></returns>
        public XrmAppointment Regarding(string entityLogicalName, Guid id)
        {
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");
            ExceptionThrow.IfGuidEmpty(id, "id");

            if (string.IsNullOrEmpty(_regardingObject.Key) && _regardingObject.Value.IsGuidEmpty())
            {
                _regardingObject = new KeyValuePair<string, Guid>(entityLogicalName.ToLower().Trim(), id);
            }

            return this;
        }

        /// <summary>
        /// Add additional attribute
        /// </summary>
        /// <param name="attributeName">Attribute name</param>
        /// <param name="value">
        /// Attribute value by MS CRM datatypes (<see cref="String"/>, <see cref="int"/> or <see cref="EntityReference"/>, <see cref="OptionSetValue"/>)
        /// </param>
        /// <returns>
        /// <see cref="XrmAppointment"/>
        /// </returns>
        public XrmAppointment AddCustomAttribute(string attributeName, object value)
        {
            ExceptionThrow.IfNullOrEmpty(attributeName, "attributeName");

            _customAttributeList.Add(attributeName, value);

            return this;
        }

        /// <summary>
        /// Convert <see cref="XrmAppointment"/> to <see cref="Entity"/>
        /// </summary>
        /// <returns><see cref="Entity"/></returns>
        public Entity ToEntity()
        {
            Entity result = null;

            if (!_isInit)
            {
                throw new InvalidOperationException("Please INIT (call constructor of XrmAppointment) before call ToEntity method.");
            }

            ExceptionThrow.IfGuidEmpty(_organizer.Value, "Organizer.Id");
            ExceptionThrow.IfNullOrEmpty(_attendeeList, "AttendeeList");
            ExceptionThrow.IfNullOrEmpty(_subject, "Subject");
            ExceptionThrow.IfNull(_start, "StartDate");
            ExceptionThrow.IfNull(_end, "EndDate");

            result = new Entity("appointment");

            Entity organizer = new Entity("activityparty");
            organizer["partyid"] = new EntityReference(_organizer.Key.Description(), _organizer.Value);

            result["subject"] = _subject;
            result["description"] = _description;
            result["organizer"] = new[] { organizer };
            result["requiredattendees"] = CreateActivityParty(AttendeeTypeCode.Required, _attendeeList);
            result["optionalattendees"] = CreateActivityParty(AttendeeTypeCode.Optional, _attendeeList);
            result["prioritycode"] = new OptionSetValue(_priority);
            result["scheduledstart"] = _start.Value;
            result["scheduledend"] = _end.Value;
            result["isalldayevent"] = _isAllDayEvent;
            result["location"] = _location;

            if (!string.IsNullOrEmpty(_regardingObject.Key) && !_regardingObject.Value.IsGuidEmpty())
            {
                result["regardingobjectid"] = new EntityReference(_regardingObject.Key, _regardingObject.Value);
            }

            if(_customAttributeList !=null && _customAttributeList.Keys.Count > 0)
            {
                foreach (KeyValuePair<string, object> item in _customAttributeList)
                {
                    result[item.Key] = item.Value;
                }
            }

            return result;
        }

        #endregion

        #region | Private Methods |

        EntityCollection CreateActivityParty(AttendeeTypeCode typeCode, List<Tuple<AttendeeTypeCode, AttendeeEntityTypeCode, Guid>> list)
        {
            EntityCollection result = null;

            if (!list.IsNullOrEmpty())
            {
                result = new EntityCollection();

                var attendeeList = list.Where(d => d.Item1.Equals(typeCode)).ToList();

                foreach (var item in attendeeList)
                {
                    Entity p = new Entity("activityparty");

                    if (!item.Item3.IsGuidEmpty())
                    {
                        p["partyid"] = new EntityReference(item.Item2.Description(), item.Item3);
                    }

                    result.Entities.Add(p);
                }
            }

            return result;
        }

        #endregion
    }
}
