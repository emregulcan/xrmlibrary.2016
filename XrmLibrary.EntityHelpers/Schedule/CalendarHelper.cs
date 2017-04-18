using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Administrator;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Schedule
{
    /// <summary>
    /// The <c>Calendar</c> entity stores data for customer service calendars and holiday schedules in addition to business. Each calendar is set for a specific time zone.
    /// This class provides mostly used common methods for <c>Calendar</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328538(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class CalendarHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Calendar.GetBusinessClosureDates
         * We can not perform filter query on Calendar / Business Closure Dates.
         * 
         */

        #endregion

        #region | Private Definitions |

        bool _useUtc = false;
        OrganizationHelper _organizationHelper;

        #endregion

        #region | Enums |

        /// <summary>
        /// Date filters
        /// </summary>
        public enum DateFilter
        {
            Today,
            Tomorrow,
            Yesterday,
            OnDate,
            OnOrAfter,
            OnOrBefore,
            OnYear,

            ThisWeek,
            ThisMonth,
            ThisYear,

            NextWeek,
            NextMonth,
            NextYear,

            NextXDays,
            NextXWeeks,
            NextXMonths,
            NextXYears,

            LastWeek,
            LastMonth,
            LastYear,

            LastXDays,
            LastXWeeks,
            LastXMonths,
            LastXYears
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        /// <param name="useUTC">
        /// Set <c>true</c>, if you want to use Utc datetime format, otherwise set <c>false</c>.
        /// This parameter's default value is <c>true</c>
        /// </param>
        public CalendarHelper(IOrganizationService service, bool useUTC = true) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328538(v=crm.8).aspx";
            this.EntityName = "calendar";

            this._useUtc = useUTC;
            this._organizationHelper = new OrganizationHelper(service);
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve all <c>Business Closure Dates</c> of <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Organization</c>
        /// </summary>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates()
        {
            Guid organizationId = this._organizationHelper.GetOrganizationId();
            return GetBusinessClosureDates(organizationId);
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Organization</c> by specified filter.
        /// </summary>
        /// <param name="filter">
        /// <see cref="DateFilter"/> Date filter
        /// </param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(DateFilter filter)
        {
            Guid organizationId = this._organizationHelper.GetOrganizationId();
            return GetBusinessClosureDates(organizationId, filter);
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Organization</c> by specified filter.
        /// </summary>
        /// <param name="filter"><see cref="DateFilter"/> Date filter</param>
        /// <param name="date"><see cref="DateTime"/> Date</param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(DateFilter filter, DateTime date)
        {
            Guid organizationId = this._organizationHelper.GetOrganizationId();
            return GetBusinessClosureDates(organizationId, filter, date);
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Organization</c> by specified filter.
        /// </summary>
        /// <param name="filter"><see cref="DateFilter"/> Date filter</param>
        /// <param name="value"><see cref="int"/> Date filter value</param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(DateFilter filter, int value)
        {
            Guid organizationId = this._organizationHelper.GetOrganizationId();
            return GetBusinessClosureDates(organizationId, filter, value);
        }

        /// <summary>
        /// Retrieve all <c>Business Closure Dates</c> of <c>Organization</c>.
        /// </summary>
        /// <param name="organizationId"><c>Organization</c> Id</param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(Guid organizationId)
        {
            ExceptionThrow.IfGuidEmpty(organizationId, "organizationId");

            var organization = this._organizationHelper.Get(organizationId, "businessclosurecalendarid");

            ExceptionThrow.IfNull(organization, "Organization", string.Format("Organization data not found related with '{0}' id", organizationId.ToString()));

            var query = new QueryExpression("calendar")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression()
            };

            query.Criteria.AddCondition(new ConditionExpression("calendarid", ConditionOperator.Equal, organization["businessclosurecalendarid"].ToString()));

            var serviceResponse = this.OrganizationService.RetrieveMultiple(query);
            return serviceResponse[0].GetAttributeValue<EntityCollection>("calendarrules");
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> by specified filter.
        /// </summary>
        /// <param name="organizationId"><c>Organization</c> Id</param>
        /// <param name="filter"><see cref="DateFilter"/> Date filter</param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(Guid organizationId, DateFilter filter)
        {
            EntityCollection bcdList = GetBusinessClosureDates(organizationId);
            return Filter(bcdList, filter, null, null);
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> by specified filter.
        /// </summary>
        /// <param name="organizationId"><c>Organization</c> Id</param>
        /// <param name="filter"><see cref="DateFilter"/> Date filter</param>
        /// <param name="date"></param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(Guid organizationId, DateFilter filter, DateTime date)
        {
            EntityCollection bcdList = GetBusinessClosureDates(organizationId);
            return Filter(bcdList, filter, date, null);
        }

        /// <summary>
        /// Retrieve all active and continued <c>Business Closure Dates</c> by specified filter.
        /// </summary>
        /// <param name="organizationId"><c>Organization</c> Id</param>
        /// <param name="filter"><see cref="DateFilter"/> Date filter</param>
        /// <param name="value"></param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Business Closure Dates</c> data
        /// </returns>
        public EntityCollection GetBusinessClosureDates(Guid organizationId, DateFilter filter, int value)
        {
            EntityCollection bcdList = GetBusinessClosureDates(organizationId);
            return Filter(bcdList, filter, null, value);
        }

        /// <summary>
        /// Search the specified resource for an available time block that matches the specified parameters.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.queryschedulerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="resourceId">Resource Id</param>
        /// <param name="start">Start of the time slot.</param>
        /// <param name="end">End of the time slot.</param>
        /// <param name="timecode">
        /// <see cref="TimeCode"/>
        /// </param>
        /// <returns><see cref="QueryScheduleResponse"/></returns>
        public QueryScheduleResponse GetWorkingHours(Guid resourceId, DateTime start, DateTime end, TimeCode timecode)
        {
            ExceptionThrow.IfGuidEmpty(resourceId, "resourceId");

            QueryScheduleRequest request = new QueryScheduleRequest()
            {
                ResourceId = resourceId,
                Start = start,
                End = end,
                TimeCodes = new TimeCode[] { timecode }
            };

            return (QueryScheduleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Search multiple resources for available time block that match the specified parameters.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.querymultipleschedulesrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="resourceIdList"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="timecodes"></param>
        /// <returns><see cref="QueryMultipleSchedulesResponse"/></returns>
        public QueryMultipleSchedulesResponse GetWorkingHours(List<Guid> resourceIdList, DateTime start, DateTime end, TimeCode[] timecodes)
        {
            ExceptionThrow.IfNullOrEmpty(resourceIdList, "resourceIdList");

            QueryMultipleSchedulesRequest request = new QueryMultipleSchedulesRequest()
            {
                ResourceIds = resourceIdList.ToArray(),
                Start = start,
                End = end,
                TimeCodes = timecodes
            };

            return (QueryMultipleSchedulesResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Filter <c>Business Closure Dates</c> data by specified data
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filter"></param>
        /// <param name="date"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        EntityCollection Filter(EntityCollection data, DateFilter filter, DateTime? date, int? value)
        {
            EntityCollection result = new EntityCollection();

            IEnumerable<Entity> filteredData = null;
            DateTime[] dateRange = new DateTime[] { };
            DateTime now = this._useUtc ? DateTime.UtcNow : DateTime.Now;
            DateTime filterDate = DateTime.MinValue;

            switch (filter)
            {
                case DateFilter.Today:
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= DateTime.Today.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= DateTime.Today.Date);
                    break;

                case DateFilter.Tomorrow:
                    filterDate = DateTime.Today.AddDays(1);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;

                case DateFilter.Yesterday:
                    filterDate = DateTime.Today.AddDays(-1);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;

                case DateFilter.OnDate:
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= date.Value.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= date.Value.Date);
                    break;

                case DateFilter.OnOrAfter:
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= date.Value.Date);
                    break;

                case DateFilter.OnOrBefore:
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date == date.Value.Date || ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) <= date.Value.Date);
                    break;

                case DateFilter.OnYear:
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date.Year == value.Value || ((DateTime)d["effectiveintervalend"]).Date.Year == value.Value);
                    break;

                case DateFilter.ThisWeek:
                    dateRange = DateTime.Today.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Weekly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.ThisMonth:
                    dateRange = DateTime.Today.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Monthly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.ThisYear:
                    dateRange = DateTime.Today.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Annually);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.NextWeek:
                    dateRange = DateTime.Today.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Weekly);
                    filterDate = dateRange[0].AddDays(7);
                    dateRange = filterDate.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Weekly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.NextMonth:
                    filterDate = DateTime.Today.AddMonths(1);
                    dateRange = filterDate.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Monthly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.NextYear:
                    filterDate = DateTime.Today.AddYears(1);
                    dateRange = filterDate.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Annually);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.NextXDays:
                    filterDate = DateTime.Today.AddDays(value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= now.Date);
                    break;

                case DateFilter.NextXWeeks:
                    filterDate = DateTime.Today.AddWeeks(value.Value).LastDayOfWeek();
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= now.Date);
                    break;

                case DateFilter.NextXMonths:
                    filterDate = DateTime.Today.AddMonths(value.Value).LastDayOfMonth();
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= now.Date);
                    break;

                case DateFilter.NextXYears:
                    filterDate = DateTime.Today.AddYears(value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= now.Date);
                    break;

                case DateFilter.LastWeek:
                    dateRange = DateTime.Today.FirstDayOfWeek().AddDays(-7).GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Weekly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.LastMonth:
                    filterDate = DateTime.Today.AddMonths(-1);
                    dateRange = filterDate.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Monthly);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.LastYear:
                    filterDate = DateTime.Today.AddYears(-1);
                    dateRange = filterDate.GetBeginAndEndDates(DateTimeExtensions.FrequencyType.Annually);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= dateRange[1].Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= dateRange[0].Date);
                    break;

                case DateFilter.LastXDays:
                    filterDate = DateTime.Today.AddDays(-value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= filterDate.Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;

                case DateFilter.LastXWeeks:
                    filterDate = DateTime.Today.AddWeeks(-value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= now.FirstDayOfWeek().Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;

                case DateFilter.LastXMonths:
                    filterDate = DateTime.Today.AddMonths(-value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= now.FirstDayOfWeek().Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;

                case DateFilter.LastXYears:
                    filterDate = DateTime.Today.AddYears(-value.Value);
                    filteredData = data.Entities.Where(d => ((DateTime)d["effectiveintervalstart"]).Date <= now.FirstDayOfWeek().Date && ((DateTime)d["effectiveintervalend"]).Date.AddDays(-1) >= filterDate.Date);
                    break;
            }

            if (filteredData != null)
            {
                result = new EntityCollection(filteredData.ToList());
            }

            return result;
        }

        #endregion
    }
}