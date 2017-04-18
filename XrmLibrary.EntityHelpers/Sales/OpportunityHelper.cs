using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// The <c>Opportunity</c> entity represents a potential sale to new or established customers. It helps you to forecast future business demands and sales revenues. 
    /// This class provides mostly used common methods for <c>Opportunity</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334362(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class OpportunityHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Opportunity 
         * 
         */

        #endregion

        #region | Private Definitions |

        bool _useUtc = false;

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Opportunity</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Opportunity
        /// </para>
        /// </summary>
        public enum OpportunityActiveStatusCode
        {
            CustomStatusCode = 0,
            InProgrress = 1,
            OnHold = 2
        }

        /// <summary>
        /// <c>Opportunity</c> 's <c>Won</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Opportunity
        /// </para>
        /// </summary>
        public enum OpportunityWonStatusCode
        {
            CustomStatusCode = 0,
            Won = 3
        }

        /// <summary>
        /// <c>Opportunity</c> 's <c>Lost</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Opportunity
        /// </para>
        /// </summary>
        public enum OpportunityLostStatusCode
        {
            CustomStatusCode = 0,
            Cancelled = 4,
            OutSold = 5
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
        public OpportunityHelper(IOrganizationService service, bool useUTC = true) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334362(v=crm.8).aspx";
            this.EntityName = "opportunity";
            this._useUtc = useUTC;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// <c>Win</c> an opportunity.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.winopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Opportunity</c> Id</param>
        /// <param name="status"><see cref="OpportunityWonStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="opportunitycloseSubject"><c>OpportunityClose</c> subjec</param>
        /// <param name="opportunitycloseDescription"><c>OpportunityClose</c> description</param>
        /// <param name="opportunitycloseActualEnd"><c>OpportunityClose</c> actual end date. Default value is <c>DateTime.UtcNow</c> or <c>DateTime.Now</c> depend on <c>useUtc</c> parameter on constructor</param>
        /// <param name="useOpportunityTotalAmountValue">Set <c>true</c>, if you want to use <c>Opportunity</c>'s current <c>Total Amount</c> value, otherwise if you want new value set <c>false</c> and set value of <c>opportunitycloseActualRevenue</c> parameter</param>
        /// <param name="opportunitycloseActualRevenue"><c>OpportunityClose</c> actual revenue.</param>
        /// <returns><see cref="WinOpportunityResponse"/></returns>
        public WinOpportunityResponse Win(Guid id, OpportunityWonStatusCode status, int customStatusCode = 0, string opportunitycloseSubject = "", string opportunitycloseDescription = "", DateTime? opportunitycloseActualEnd = null, bool useOpportunityTotalAmountValue = false, decimal opportunitycloseActualRevenue = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;
            opportunitycloseSubject = !string.IsNullOrEmpty(opportunitycloseSubject) ? opportunitycloseSubject.Trim() : "Opportunity Close - Won";

            if (status == OpportunityWonStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            if (useOpportunityTotalAmountValue)
            {
                opportunitycloseActualRevenue = GetTotalAmount(id);
            }

            WinOpportunityRequest request = new WinOpportunityRequest()
            {
                OpportunityClose = PrepareOpportunityCloseEntity(id, opportunitycloseSubject, opportunitycloseDescription, null, opportunitycloseActualEnd, opportunitycloseActualRevenue),
                Status = new OptionSetValue(statusCode)
            };

            return (WinOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Close an <c>opportunity</c> as <c>Lost</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.loseopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Opportunity</c> Id</param>
        /// <param name="competitorId">Competitor</param>
        /// <param name="status"><see cref="OpportunityLostStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="opportunitycloseSubject"><c>OpportunityClose</c> subject</param>
        /// <param name="opportunitycloseDescription"><c>OpportunityClose</c> description</param>
        /// <param name="opportunitycloseActualEnd"><c>OpportunityClose</c> actual end date. Default value is <c>DateTime.UtcNow</c> or <c>DateTime.Now</c> depend on <c>useUtc</c> parameter on constructor</param>
        /// <param name="opportunitycloseActualRevenue"><c>OpportunityClose</c> actual revenue. Mostly this parameter's value is 0 (zero) when you lost an opportunity</param>
        /// <returns><see cref="LoseOpportunityResponse"/></returns>
        public LoseOpportunityResponse Lose(Guid id, Guid? competitorId, OpportunityLostStatusCode status, int customStatusCode = 0, string opportunitycloseSubject = "", string opportunitycloseDescription = "", DateTime? opportunitycloseActualEnd = null, decimal opportunitycloseActualRevenue = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;
            opportunitycloseSubject = !string.IsNullOrEmpty(opportunitycloseSubject) ? opportunitycloseSubject.Trim() : "Opportunity Close - Lost";

            if (status == OpportunityLostStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            LoseOpportunityRequest request = new LoseOpportunityRequest()
            {
                OpportunityClose = PrepareOpportunityCloseEntity(id, opportunitycloseSubject, opportunitycloseDescription, competitorId, opportunitycloseActualEnd, opportunitycloseActualRevenue),
                Status = new OptionSetValue(statusCode)
            };

            return (LoseOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Calculate the value of an <c>Opportunity</c> that is in the <c>Won</c> state.
        /// <para>
        /// Please note that this method works correctly if opportunity has valid <c>Quote</c> records that <c>Won</c> state (converted to <c>SalesOrder</c>).
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.calculateactualvalueopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Opportunity</c> Id</param>
        /// <returns>
        /// <see cref="CalculateActualValueOpportunityResponse.Value"/> attribute for calculated revenue
        /// </returns>
        public CalculateActualValueOpportunityResponse CalculateActualValue(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CalculateActualValueOpportunityRequest request = new CalculateActualValueOpportunityRequest()
            {
                OpportunityId = id
            };

            return (CalculateActualValueOpportunityResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// 
        /// </summary>
        /// <param name="opportunityId"></param>
        /// <returns></returns>
        decimal GetTotalAmount(Guid opportunityId)
        {
            decimal result = 0;
            var entity = this.Get(opportunityId, "totalamount");
            result = entity.GetAttributeValue<Money>("totalamount").Value;

            return result;
        }

        /// <summary>
        /// Create an <c>opportunityclose</c> entity
        /// </summary>
        /// <param name="opportunityId">Existing Opportunity record Id</param>
        /// <param name="subject">Subject</param>
        /// <param name="description">Description</param>
        /// <param name="competitorId">Competitor</param>
        /// <param name="actualEnd">Actual End date</param>
        /// <param name="actualRevenue">Actual Revenue</param>
        /// <returns></returns>
        Entity PrepareOpportunityCloseEntity(Guid opportunityId, string subject, string description, Guid? competitorId, DateTime? actualEnd = null, decimal actualRevenue = 0)
        {
            DateTime now = this._useUtc ? DateTime.UtcNow : DateTime.Now;

            Entity opportunityClose = new Entity("opportunityclose");
            opportunityClose["opportunityid"] = new EntityReference("opportunity", opportunityId);
            opportunityClose["subject"] = subject;
            opportunityClose["description"] = description;
            opportunityClose["actualend"] = actualEnd.HasValue ? actualEnd.Value : now;
            opportunityClose["actualrevenue"] = new Money(actualRevenue);

            if (competitorId.HasValue && !competitorId.Value.IsGuidEmpty())
            {
                opportunityClose["competitorid"] = new EntityReference("competitor", competitorId.Value);
            }

            return opportunityClose;
        }

        #endregion
    }
}
