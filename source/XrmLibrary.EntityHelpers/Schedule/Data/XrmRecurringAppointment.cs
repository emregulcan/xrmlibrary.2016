using System;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;

namespace XrmLibrary.EntityHelpers.Schedule.Data
{
    /// <summary>
    /// Fluent object to easy-create MS CRM <c>Recurring Appointment</c> activity.
    /// </summary>
    public class XrmRecurringAppointment
    {
        #region | Private Definitions |

        string _hasRecurrenceException = "You can only add a recurrence.";
        string _hasRecurrenceEndException = "You can only add a recurrence end option.";

        bool _isInit = false;
        bool _hasRecurrence = false;
        bool _hasEndPattern = false;
        bool _isEveryWeekday = false;
        bool _isFirstOption = false;

        DateTime _startTime = DateTime.MinValue;
        DateTime _endTime = DateTime.MinValue;
        DateTime _startRangeDate = DateTime.MinValue;
        DateTime _endRangeDate = DateTime.MinValue;

        int _dayOfWeeksValue = 0;
        int _interval = 0;
        int _dayNumber = 0;
        int _occurence = 0;

        RecurrencePatternType _recurrencePattern;
        RecurrenceEndPatternType _recurrenceEndPattern;
        MonthlyOption _monthlyOption;
        Month _month;

        #endregion

        #region | Enums |

        /// <summary>
        /// Recurrence Pattern
        /// </summary>
        public enum RecurrencePatternType
        {
            /// <summary>
            /// Daily recurrence
            /// </summary>
            Daily = 0,

            /// <summary>
            /// Weekly recurrence
            /// </summary>
            Weekly = 1,

            /// <summary>
            /// Monthly recurrence
            /// </summary>
            Monthly = 2,

            /// <summary>
            /// Yearly recurrence
            /// </summary>
            Yearly = 3
        }

        /// <summary>
        /// Recurrence End Pattern
        /// </summary>
        public enum RecurrenceEndPatternType
        {
            /// <summary>
            /// 
            /// </summary>
            NoEndDate = 1,

            /// <summary>
            /// 
            /// </summary>
            Occurrences = 2,

            /// <summary>
            /// 
            /// </summary>
            WithEnddate = 3
        }

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        public enum DayOfWeek
        {
            //INFO : http://numbermonk.com/

            /// <summary>
            /// 0x01 : 1
            /// </summary>
            Sunday = 0x01,

            /// <summary>
            /// 0x02 : 2
            /// </summary>
            Monday = 0x02,

            /// <summary>
            /// 0x04 : 4
            /// </summary>
            Tuesday = 0x04,

            /// <summary>
            /// 0x08 : 8
            /// </summary>
            Wednesday = 0x08,

            /// <summary>
            /// 0x10 : 16
            /// </summary>
            Thursday = 0x10,

            /// <summary>
            /// 0x20 : 32
            /// </summary>
            Friday = 0x20,

            /// <summary>
            /// 0x40 : 64
            /// </summary>
            Saturday = 0x40,

            /// <summary>
            /// 0x3e : 62
            /// </summary>
            AllWeekdays = 0x3e,

            /// <summary>
            /// 0x7f : 127
            /// </summary>
            AllDays = 0x7f
        }

        /// <summary>
        /// 
        /// </summary>
        public enum Month
        {
            Undefined = 0,
            January = 1,
            February = 2,
            March = 3,
            April = 4,
            May = 5,
            June = 6,
            July = 7,
            August = 8,
            September = 9,
            October = 10,
            November = 11,
            December = 12
        }

        /// <summary>
        /// 
        /// </summary>
        public enum MonthlyOption
        {
            First = 1,
            Second = 2,
            Third = 3,
            Fourth = 4,
            Last = 5
        }

        #endregion

        #region | Constructors |

        public XrmRecurringAppointment()
        {
            _isInit = true;
            _recurrencePattern = RecurrencePatternType.Daily;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Appointment Start Time (only time)
        /// </summary>
        /// <param name="time">
        /// Only Time format.
        /// </param>
        /// <returns></returns>
        public XrmRecurringAppointment StartTime(DateTime time)
        {
            ExceptionThrow.IfEquals(time, "StartTime", DateTime.MinValue);

            _startTime = time;
            return this;
        }

        /// <summary>
        /// Appointment End Time (only time)
        /// </summary>
        /// <param name="time">Only Time format.</param>
        /// <returns></returns>
        public XrmRecurringAppointment EndTime(DateTime time)
        {
            ExceptionThrow.IfEquals(time, "EndTime", DateTime.MinValue);

            _endTime = time;
            return this;
        }

        /// <summary>
        /// Appointment Start date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public XrmRecurringAppointment StartDate(DateTime date)
        {
            ExceptionThrow.IfEquals(date, "StartDate", DateTime.MinValue);

            _startRangeDate = date;
            return this;
        }

        /// <summary>
        /// Appointment End date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public XrmRecurringAppointment EndDate(DateTime date)
        {
            ExceptionThrow.IfEquals(date, "EndDate", DateTime.MinValue);

            _endRangeDate = date;
            return this;
        }

        /// <summary>
        /// <c>Daily</c> recurrence with <c>every weekdays</c> (Monday - Tuesday - Wednesday - Thursday - Friday).
        /// </summary>
        /// <returns></returns>
        public XrmRecurringAppointment Daily()
        {
            SetRecurrence();

            _recurrencePattern = RecurrencePatternType.Daily;
            _isEveryWeekday = true;
            _dayOfWeeksValue = (int)DayOfWeek.AllWeekdays;

            return this;
        }

        /// <summary>
        /// <c>Daily</c> recurrence with <c>every X days</c>.
        /// Please note that this includes alldays (weekdays and weekend).
        /// </summary>
        /// <param name="interval"><c>every X days</c> interval value.</param>
        /// <returns></returns>
        public XrmRecurringAppointment Daily(int interval)
        {
            ExceptionThrow.IfEquals(interval, "interval", 0);
            ExceptionThrow.IfNegative(interval, "interval");

            SetRecurrence();

            _recurrencePattern = RecurrencePatternType.Daily;
            _interval = interval;

            return this;
        }

        /// <summary>
        /// <c>Weekly</c> recurrence.
        /// </summary>
        /// <param name="interval"></param>
        /// <param name="daysOfWeek"></param>
        /// <returns></returns>
        public XrmRecurringAppointment Weekly(int interval, DayOfWeek daysOfWeek)
        {
            ExceptionThrow.IfEquals(interval, "interval", 0);
            ExceptionThrow.IfNegative(interval, "interval");

            SetRecurrence();

            _recurrencePattern = RecurrencePatternType.Weekly;
            _interval = interval;
            _dayOfWeeksValue = (int)daysOfWeek;

            return this;
        }

        /// <summary>
        /// <c>Monthly</c> recurrence with exact day.
        /// </summary>
        /// <param name="onDay">Excat day number</param>
        /// <param name="interval"><c>Every X months</c> interval value.</param>
        /// <returns></returns>
        public XrmRecurringAppointment Monthly(int onDay, int interval)
        {
            ExceptionThrow.IfEquals(onDay, "onDay", 0);
            ExceptionThrow.IfNegative(onDay, "onDay");
            ExceptionThrow.IfEquals(interval, "interval", 0);
            ExceptionThrow.IfNegative(interval, "interval");

            SetRecurrence();

            _isFirstOption = true;
            _recurrencePattern = RecurrencePatternType.Monthly;
            _interval = interval;
            _dayNumber = onDay;

            return this;
        }

        /// <summary>
        /// <c>Monthly</c> recurrence.
        /// </summary>
        /// <param name="option"></param>
        /// <param name="day"></param>
        /// <param name="interval"></param>
        /// <returns></returns>
        public XrmRecurringAppointment Monthly(MonthlyOption option, DayOfWeek day, int interval)
        {
            ExceptionThrow.IfEquals(interval, "interval", 0);
            ExceptionThrow.IfNegative(interval, "interval");

            SetRecurrence();

            _recurrencePattern = RecurrencePatternType.Monthly;
            _monthlyOption = option;
            _interval = interval;
            _dayOfWeeksValue = (int)day;

            return this;
        }

        /// <summary>
        /// <c>Yearly</c> recurrence with exact date.
        /// </summary>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <param name="interval"><c>Recur every X years</c> interval.</param>
        /// <returns></returns>
        public XrmRecurringAppointment Yearly(Month month, int day, int interval)
        {
            ExceptionThrow.IfEquals(day, "day", 0);
            ExceptionThrow.IfNegative(day, "day");

            SetRecurrence();

            _isFirstOption = true;
            _recurrencePattern = RecurrencePatternType.Yearly;
            _interval = interval;
            _dayNumber = day;
            _month = month;

            return this;
        }

        /// <summary>
        /// <c>Yearly</c> recurrence.
        /// </summary>
        /// <param name="option"></param>
        /// <param name="day"></param>
        /// <param name="month"></param>
        /// <param name="interval"><c>Recur every X years</c> interval.</param>
        /// <returns></returns>
        public XrmRecurringAppointment Yearly(MonthlyOption option, DayOfWeek day, Month month, int interval)
        {
            ExceptionThrow.IfEquals(interval, "interval", 0);
            ExceptionThrow.IfNegative(interval, "interval");

            SetRecurrence();

            _recurrencePattern = RecurrencePatternType.Yearly;
            _monthlyOption = option;
            _dayOfWeeksValue = (int)day;
            _month = month;
            _interval = interval;

            return this;
        }

        /// <summary>
        /// Recurrence with <c>no end date</c> option.
        /// </summary>
        /// <returns></returns>
        public XrmRecurringAppointment EndWithNoEndDate()
        {
            SetEndPattern();

            _recurrenceEndPattern = RecurrenceEndPatternType.NoEndDate;

            return this;
        }

        /// <summary>
        /// Recurrence with <c>occurence(s)</c>.
        /// </summary>
        /// <param name="occurence"></param>
        /// <returns></returns>
        public XrmRecurringAppointment EndAfterXOccurences(int occurence)
        {
            ExceptionThrow.IfNegative(occurence, "occurence");
            ExceptionThrow.IfEquals(occurence, "occurence", 0);

            SetEndPattern();

            _recurrenceEndPattern = RecurrenceEndPatternType.Occurrences;
            _occurence = occurence;

            return this;
        }

        /// <summary>
        /// Recurrence with <c>end date</c>.
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public XrmRecurringAppointment EndWithDate(DateTime date)
        {
            ExceptionThrow.IfEquals(date, "EndWithDate", DateTime.MinValue);

            SetEndPattern();

            _recurrenceEndPattern = RecurrenceEndPatternType.WithEnddate;
            _endRangeDate = date;

            return this;
        }

        /// <summary>
        /// Convert <see cref="XrmRecurringAppointment"/> to MS CRM <see cref="Entity"/> for <c>Recurring Appointment</c>.
        /// </summary>
        /// <returns>
        /// <see cref="Microsoft.Xrm.Sdk.Entity"/>
        /// </returns>
        public Entity ToEntity()
        {
            Entity result = null;

            if (!_isInit)
            {
                throw new InvalidOperationException("Please INIT (call constructor of XrmRecurringAppointment) before call ToEntity method.");
            }

            Validate();

            result = new Entity("recurringappointmentmaster");
            result["starttime"] = _startTime;
            result["endtime"] = _endTime;
            result["recurrencepatterntype"] = new OptionSetValue((int)_recurrencePattern);
            result["patternendtype"] = new OptionSetValue((int)_recurrenceEndPattern);
            result["patternstartdate"] = _startRangeDate;
            result["interval"] = _interval;

            switch (_recurrencePattern)
            {
                case RecurrencePatternType.Daily:

                    if (_isEveryWeekday)
                    {
                        /*
                         * INFO : Daily - every weekday pattern flow;
                         * we remove "interval" and "recurrencepatterntype (Daily)"  attributes.
                         * after that add "recurrencepatterntype" attribute with "Weekly" value.
                         * Even though "Daily - Every Weekday"  is selected from the MS Dynamics CRM UI, the input data is changed to "Weekly" pattern when inserting to database.
                         * 
                         * "Daily - Every Weekday" pattern --> "Weekly" + "isweekdaypattern = true" + daysofweekmask = 62 (all weekdays value total)
                         */

                        result.Attributes.Remove("interval");
                        result.Attributes.Remove("recurrencepatterntype");

                        result["recurrencepatterntype"] = new OptionSetValue((int)RecurrencePatternType.Weekly);
                        result["isweekdaypattern"] = true;
                        result["daysofweekmask"] = (int)DayOfWeek.AllWeekdays; //62
                    }

                    break;

                case RecurrencePatternType.Weekly:
                    result["daysofweekmask"] = _dayOfWeeksValue;

                    break;

                case RecurrencePatternType.Monthly:

                    if (_isFirstOption)
                    {
                        result["dayofmonth"] = _dayNumber;
                    }
                    else
                    {
                        result["isnthmonthly"] = true;
                        result["instance"] = new OptionSetValue((int)_monthlyOption);
                        result["daysofweekmask"] = _dayOfWeeksValue;
                    }

                    break;

                case RecurrencePatternType.Yearly:
                    result["monthofyear"] = new OptionSetValue((int)_month);

                    if (_isFirstOption)
                    {
                        result["dayofmonth"] = _dayNumber;
                    }
                    else
                    {
                        result["isnthyearly"] = true;
                        result["instance"] = new OptionSetValue((int)_monthlyOption);
                        result["daysofweekmask"] = _dayOfWeeksValue;
                    }

                    break;
            }

            switch (_recurrenceEndPattern)
            {
                case RecurrenceEndPatternType.NoEndDate:
                    break;

                case RecurrenceEndPatternType.Occurrences:
                    result["occurrences"] = _occurence;
                    break;

                case RecurrenceEndPatternType.WithEnddate:
                    result["patternenddate"] = _endRangeDate;
                    break;
            }

            return result;
        }

        /// <summary>
        /// Validate <see cref="XrmRecurringAppointment"/> required fields by patterns and throws exception if values not expected.
        /// </summary>
        public void Validate()
        {
            ExceptionThrow.IfEquals(_startTime, "StartTime", DateTime.MinValue);
            ExceptionThrow.IfEquals(_endTime, "EndTime", DateTime.MinValue);
            ExceptionThrow.IfEquals(_startRangeDate, "StartRange", DateTime.MinValue);

            #region | Validate Recurrence Pattern |

            switch (_recurrencePattern)
            {
                case RecurrencePatternType.Daily:

                    if (_isEveryWeekday)
                    {
                        ExceptionThrow.IfEquals(_dayOfWeeksValue, "Days Of Week", 0);
                        ExceptionThrow.IfNegative(_dayOfWeeksValue, "Days Of Week");
                        ExceptionThrow.IfNotExpectedValue(_dayOfWeeksValue, "Days Of Week", 62);
                    }
                    else
                    {
                        ExceptionThrow.IfEquals(_interval, "Interval", 0);
                        ExceptionThrow.IfNegative(_interval, "Interval");
                    }

                    break;

                case RecurrencePatternType.Weekly:
                    ExceptionThrow.IfEquals(_dayOfWeeksValue, "Days Of Week", 0);
                    ExceptionThrow.IfNegative(_dayOfWeeksValue, "Days Of Week");
                    ExceptionThrow.IfGreaterThan(_dayOfWeeksValue, "Days Of Week", 127);
                    ExceptionThrow.IfEquals(_interval, "Interval", 0);
                    ExceptionThrow.IfNegative(_interval, "Interval");

                    break;

                case RecurrencePatternType.Monthly:
                    ExceptionThrow.IfEquals(_interval, "Interval", 0);
                    ExceptionThrow.IfNegative(_interval, "Interval");

                    if (_isFirstOption)
                    {
                        ExceptionThrow.IfEquals(_dayNumber, "Day", 0);
                        ExceptionThrow.IfNegative(_dayNumber, "Day");
                    }
                    else
                    {
                        ExceptionThrow.IfEquals(_dayOfWeeksValue, "Days Of Week", 0);
                        ExceptionThrow.IfNegative(_dayOfWeeksValue, "Days Of Week");
                        ExceptionThrow.IfGreaterThan(_dayOfWeeksValue, "Days Of Week", 127);
                    }

                    break;

                case RecurrencePatternType.Yearly:

                    ExceptionThrow.IfEquals(_interval, "Interval", 0);
                    ExceptionThrow.IfNegative(_interval, "Interval");
                    ExceptionThrow.IfEquals((int)_month, "Month", 0);
                    ExceptionThrow.IfNegative((int)_month, "Month");

                    if (_isFirstOption)
                    {
                        ExceptionThrow.IfEquals(_dayNumber, "Day", 0);
                        ExceptionThrow.IfNegative(_dayNumber, "Day");
                    }
                    else
                    {
                        ExceptionThrow.IfEquals(_dayOfWeeksValue, "Days Of Week", 0);
                        ExceptionThrow.IfNegative(_dayOfWeeksValue, "Days Of Week");
                        ExceptionThrow.IfGreaterThan(_dayOfWeeksValue, "Days Of Week", 127);
                    }

                    break;
            }

            #endregion

            #region | Validate Recurrence End Pattern |

            switch (_recurrenceEndPattern)
            {
                case RecurrenceEndPatternType.NoEndDate:
                    break;

                case RecurrenceEndPatternType.Occurrences:
                    ExceptionThrow.IfNegative(_occurence, "Occurence");
                    ExceptionThrow.IfEquals(_occurence, "Occurence", 0);

                    break;

                case RecurrenceEndPatternType.WithEnddate:
                    ExceptionThrow.IfEquals(_endRangeDate, "EndRange", DateTime.MinValue);

                    break;
            }

            #endregion
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Check and set recurrence data.
        /// If has any recurrence throws exception.
        /// </summary>
        void SetRecurrence()
        {
            if (_hasRecurrence)
            {
                throw new ArgumentException(_hasRecurrenceException);
            }

            _hasRecurrence = true;
        }

        /// <summary>
        /// Check and set recurrence end pattern data.
        /// If has any recurrence throws exception.
        /// </summary>
        void SetEndPattern()
        {
            if (_hasEndPattern)
            {
                throw new ArgumentException(_hasRecurrenceEndException);
            }

            _hasEndPattern = true;
        }

        #endregion
    }
}
