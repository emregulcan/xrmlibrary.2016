using System;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Service
{
    /// <summary>
    /// An <c>Incident (case)</c> is a customer service issue, which is a problem that is defined by a customer and the collection of activities that occur in the case to resolve the problem.
    /// All actions and communications can be tracked in the <c>Incident</c> entity. 
    /// This class provides mostly used common methods for <c>Incident</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328291(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class IncidentHelper : BaseEntityHelper
    {
        #region | Enums |

        /// <summary>
        /// <c>Incident</c> 's statecode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Case
        /// </para>
        /// </summary>
        public enum IncidentStateCode
        {
            [Description("active")]
            Active = 0,

            [Description("resolved")]
            Resolved = 1,

            [Description("calceled")]
            Canceled = 2
        }

        /// <summary>
        /// <c>Incident</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Case
        /// </para>
        /// </summary>
        public enum IncidentActiveStatusCode
        {
            CustomStatusCode = 0,
            InProgress = 1,
            OnHold = 2,
            WaitingForDetails = 3,
            Researching = 4
        }

        /// <summary>
        /// <c>Incident</c> 's <c>Resolved</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Case
        /// </para>
        /// </summary>
        public enum IncidentResolvedStatusCode
        {
            CustomStatusCode = 0,
            Solved = 5,
            InformationProvided = 1000
        }

        /// <summary>
        /// <c>Incident</c> 's <c>Canceled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Case
        /// </para>
        /// </summary>
        public enum IncidentCanceledStatusCode
        {
            CustomStatusCode = 0,
            Canceled = 6,
            Merged = 2000
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public IncidentHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328291(v=crm.8).aspx";
            this.EntityName = "incident";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Applies the active <c>Routing Rule</c> to the <c>Incident (case)</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.applyroutingrulerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Incident</c> Id</param>
        /// <returns>
        /// <see cref="ApplyRoutingRuleResponse"/>
        /// </returns>
        public ApplyRoutingRuleResponse ApplyRoute(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            ApplyRoutingRuleRequest request = new ApplyRoutingRuleRequest()
            {
                Target = new EntityReference(this.EntityName, id)
            };

            return (ApplyRoutingRuleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Calculates the total time that was spent on the <c>Incident</c>.
        /// Please note that if an activity that is used to resolve the incident (case) has already been <c>billed</c>, those minutes are not included in the calculation of the current incident.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.calculatetotaltimeincidentrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Incident</c> Id</param>
        /// <returns>
        /// Calculated total time in <c>minutes</c>
        /// </returns>
        public long CalculateTotalTime(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CalculateTotalTimeIncidentRequest request = new CalculateTotalTimeIncidentRequest()
            {
                IncidentId = id
            };

            var serviceResponse = (CalculateTotalTimeIncidentResponse)this.OrganizationService.Execute(request);
            return serviceResponse.TotalTime;
        }

        /// <summary>
        /// Merge two <c>Incident</c> records.
        /// Please note that the behavior of merge for incidents is different from merging accounts, contacts, or leads.
        /// <para>
        /// For more information please look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mergeToRecordId"></param>
        /// <param name="mergedRecordId"></param>
        /// <param name="performParentingChecks"></param>
        /// <returns>
        /// <see cref="MergeResponse"/>
        /// </returns>
        public MergeResponse Merge(Guid mergeToRecordId, Guid mergedRecordId, bool performParentingChecks)
        {
            ExceptionThrow.IfGuidEmpty(mergeToRecordId, "mergeToRecordId");
            ExceptionThrow.IfGuidEmpty(mergedRecordId, "mergedRecordId");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            return commonHelper.Merge(CommonHelper.MergeEntity.Incident, mergeToRecordId, mergedRecordId, new Entity(this.EntityName), performParentingChecks);

            /*
             * INFO : Merge incidents
             *     The behavior of merge for incidents is different from merging accounts, contacts, or leads in the following ways:
             *     
             *     ### The "UpdateContent" property is not used. 
             *         For other entities this property may be used to transfer selected values from the subordinate record to the target record. 
             *         When merging incidents the value of this property is ignored.
             *     
             *     ### Merge is performed in the security context of the user.
             *         Merge operations for other entities are performed with a system user security context. 
             *         Because incident merge operations are performed in the security context of the user, the user must have the security privileges to perform any of the actions, such as re-parenting related records, that are performed by the merge operation.
             *         If the user merging records doesn’t have privileges for all the actions contained within the merge operation, the merge operation will fail and roll back to the original state.
             */
        }

        /// <summary>
        /// Re-activate an <c>Incident</c>.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="status"><see cref="IncidentActiveStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        public void Activate(Guid id, IncidentActiveStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == IncidentActiveStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)IncidentStateCode.Active, statusCode);
        }

        /// <summary>
        /// Successfuly <c>close (resolve)</c> an incident.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.closeincidentrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Incident</c> Id</param>
        /// <param name="status"><see cref="IncidentResolvedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="subject"><c>Incident Resolution</c> subject</param>
        /// <param name="description"><c>Incident Resolution</c> description</param>
        /// <param name="resolveDate"><c>Incident Resolution</c> acual end date</param>
        /// <returns>
        /// <see cref="CloseIncidentResponse"/>
        /// </returns>
        public CloseIncidentResponse Resolve(Guid id, IncidentResolvedStatusCode status, int customStatusCode = 0, string subject = "Resolved Incident", string description = "", DateTime? resolveDate = null)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == IncidentResolvedStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            Entity incidentResolution = new Entity("incidentresolution");
            incidentResolution["incidentid"] = new EntityReference(this.EntityName, id);
            incidentResolution["subject"] = subject;
            incidentResolution["description"] = description;

            if (resolveDate.HasValue && resolveDate.Value != DateTime.MinValue)
            {
                incidentResolution["actualend"] = resolveDate.Value;
            }

            CloseIncidentRequest request = new CloseIncidentRequest()
            {
                IncidentResolution = incidentResolution,
                Status = new OptionSetValue(statusCode)
            };

            return (CloseIncidentResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// <c>Cancel</c> an <c>Incident</c>.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="status"><see cref="IncidentCanceledStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        public void Cancel(Guid id, IncidentCanceledStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == IncidentCanceledStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)IncidentStateCode.Canceled, statusCode);
        }

        #endregion
    }
}
