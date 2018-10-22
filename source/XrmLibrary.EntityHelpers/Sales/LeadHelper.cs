using System;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.BusinessManagement;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// A <c>Lead</c> entity represents an individual that is identified as someone who is interested in receiving specific information about the products or services offered by the company. 
    /// A lead is used to track contacts or accounts that are potential customers, but who have not yet been qualified.
    /// This class provides mostly used common methods for <c>Lead</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328442(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class LeadHelper : BaseEntityHelper
    {
        #region | Bookmarks |

        /*
         * INFO : Bookmarks.Lead
         * 
         */

        #endregion

        #region | Enums |

        [Flags]
        public enum QualifyLeadTo
        {
            None = 1,
            Account = 2,
            Contact = 4,
            /// <summary>
            /// The created account becomes a parent (Contact.Parentcustomerid) of the contact and the contact becomes a primary contact (Account.PrimaryContactId) for the account
            /// </summary>
            AccountAndContact = 8,
            OpportunityWithContact = 16,
            OpportunityWithExistingRecord = 32,
            OpportunityWithAccountAndContact = 64
        }

        public enum ExistingEntityType
        {
            [Description("account")]
            Account = 1,

            [Description("contact")]
            Contact = 2
        }

        /// <summary>
        /// <c>Lead</c> 's <c>Open</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Lead
        /// </para>
        /// </summary>
        public enum LeadOpenStatusCode
        {
            CustomStatusCode = 0,
            New = 1,
            Contacted = 2
        }

        /// <summary>
        /// <c>Lead</c> 's <c>Qualified</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Lead
        /// </para>
        /// </summary>
        public enum LeadQualifiedStatusCode
        {
            CustomStatusCode = 0,
            Qualified = 3
        }

        /// <summary>
        /// <c>Lead</c> 's <c>Disqualified</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Lead
        /// </para>
        /// </summary>
        public enum LeadDisqualifiedStatusCode
        {
            CustomStatusCode = 0,
            Lost = 4,
            CannotContact = 5,
            NoLongerInterested = 6,
            Canceled = 7
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service">
        /// <see cref="Microsoft.Xrm.Sdk.IOrganizationService"/>
        /// </param>
        public LeadHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328442(v=crm.8).aspx";
            this.EntityName = "lead";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// <c>Qualify</c> a <c>Lead</c> and create an account, contact, or opportunity records that are linked to the originating lead.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.qualifyleadrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Lead</c> Id</param>
        /// <param name="qualifyTo">Lead qualified to entity type</param>
        /// <param name="status"><see cref="LeadQualifiedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>Resturns created entities after qualification in <see cref="QualifyLeadResponse.CreatedEntities"/> property.</returns>
        public QualifyLeadResponse Qualify(Guid id, QualifyLeadTo qualifyTo, LeadQualifiedStatusCode status = LeadQualifiedStatusCode.Qualified, int customStatusCode = 0)
        {
            return Qualify(id, qualifyTo, null, null, null, null, status, customStatusCode);
        }

        /// <summary>
        /// <c>Qualify</c> a <c>Lead</c> with existing record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.qualifyleadrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Lead</c> Id</param>
        /// <param name="qualifyTo">Lead qualified to entity type</param>
        /// <param name="existingRecordType">Existing record entity type</param>
        /// <param name="existingRecordId">If you qualify with existing record (account or contact) set record Id</param>
        /// <param name="status"><see cref="LeadQualifiedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// Resturns created entities after qualification in <see cref="QualifyLeadResponse.CreatedEntities"/> property.
        /// </returns>
        public QualifyLeadResponse Qualify(Guid id, QualifyLeadTo qualifyTo, ExistingEntityType existingRecordType, Guid existingRecordId, LeadQualifiedStatusCode status = LeadQualifiedStatusCode.Qualified, int customStatusCode = 0)
        {
            return Qualify(id, qualifyTo, existingRecordType, existingRecordId, null, null, status, customStatusCode);
        }

        /// <summary>
        /// <c>Qualify</c> a <c>Lead</c> and create an account, contact, or opportunity records that are linked to the originating lead.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.qualifyleadrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Lead</c> Id</param>
        /// <param name="qualifyTo">Lead qualified to entity type</param>
        /// <param name="existingRecordType">Existing record entity type</param>
        /// <param name="existingRecordId">If you qualify with existing record (account or contact) set record Id</param>
        /// <param name="currencyId">
        /// If you want create an <c>Opportunity</c> set <c>TransactionCurrency Id</c>, otherwise leave NULL.
        /// <para>
        /// If you want create an <c>Opportunity</c> and do not know / sure <c>TransactionCurrency Id</c> leave NULL, this method will find <c>TransactionCurrency Id</c>
        /// </para>
        /// </param>
        /// <param name="campaignId"></param>
        /// <param name="status"><see cref="LeadQualifiedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>Resturns created entities after qualification in <see cref="QualifyLeadResponse.CreatedEntities"/> property.</returns>
        public QualifyLeadResponse Qualify(Guid id, QualifyLeadTo qualifyTo, ExistingEntityType? existingRecordType, Guid? existingRecordId, Guid? currencyId, Guid? campaignId, LeadQualifiedStatusCode status = LeadQualifiedStatusCode.Qualified, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNegative((int)qualifyTo, "qualifyTo");

            int statusCode = (int)status;

            if (status == LeadQualifiedStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            bool isAccount = (qualifyTo.HasFlag(QualifyLeadTo.Account) || qualifyTo.HasFlag(QualifyLeadTo.AccountAndContact) || qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithAccountAndContact)) && !qualifyTo.HasFlag(QualifyLeadTo.None);
            bool isContact = (qualifyTo.HasFlag(QualifyLeadTo.Contact) || qualifyTo.HasFlag(QualifyLeadTo.AccountAndContact) || qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithContact) || qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithAccountAndContact)) && !qualifyTo.HasFlag(QualifyLeadTo.None);
            bool isOpportunity = (qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithContact) || qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithAccountAndContact) || qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithExistingRecord)) && !qualifyTo.HasFlag(QualifyLeadTo.None);

            QualifyLeadRequest request = new QualifyLeadRequest()
            {
                LeadId = new EntityReference(this.EntityName.ToLower(), id),
                CreateAccount = isAccount,
                CreateContact = isContact,
                CreateOpportunity = isOpportunity,
                Status = new OptionSetValue(statusCode)
            };

            if (campaignId.HasValue)
            {
                ExceptionThrow.IfGuid(campaignId.Value, "campaignId");

                request.SourceCampaignId = new EntityReference("campaign", campaignId.Value);
            }

            if (isOpportunity)
            {
                var currencyEntityReference = currencyId.HasValue && !currencyId.Value.IsGuidEmpty() ? new EntityReference("transactioncurrency", currencyId.Value) : GetTransactionCurrency("lead", id);

                ExceptionThrow.IfNull(currencyEntityReference, "OpportunityCurrency");
                ExceptionThrow.IfGuidEmpty(currencyEntityReference.Id, "OpportunityCurrency.Id");

                request.OpportunityCurrencyId = currencyEntityReference;

                if (qualifyTo.HasFlag(QualifyLeadTo.OpportunityWithExistingRecord))
                {
                    ExceptionThrow.IfNull(existingRecordType, "existingRecordType");
                    ExceptionThrow.IfGuidEmpty(existingRecordId.Value, "existingRecordId");

                    request.OpportunityCustomerId = new EntityReference(existingRecordType.Description(), existingRecordId.Value);
                }
            }

            return (QualifyLeadResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// <c>Disqualify</c> a <c>Lead</c> and set state to passive.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setstaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Lead</c> Id</param>
        /// <param name="status"><see cref="LeadDisqualifiedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        public void Disqualify(Guid id, LeadDisqualifiedStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == LeadDisqualifiedStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper entityHelper = new CommonHelper(this.OrganizationService);
            entityHelper.UpdateState(id, this.EntityName, 2, statusCode);
        }

        /// <summary>
        /// Merge two leads.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mergeToRecordId"></param>
        /// <param name="mergedRecordId"></param>
        /// <param name="updatedContent">Additional entity attributes to be set during the merge operation</param>
        /// <param name="performParentingChecks">Indicates whether to check if the parent information is different for the two entity records.</param>
        /// <returns><see cref="MergeResponse"/></returns>
        public MergeResponse Merge(Guid mergeToRecordId, Guid mergedRecordId, Entity updatedContent, bool performParentingChecks = false)
        {
            ExceptionThrow.IfGuidEmpty(mergeToRecordId, "mergeToRecordId");
            ExceptionThrow.IfGuidEmpty(mergedRecordId, "mergedRecordId");
            ExceptionThrow.IfNull(updatedContent, "updatedContent");

            CommonHelper entityHelper = new CommonHelper(this.OrganizationService);
            return entityHelper.Merge(CommonHelper.MergeEntity.Lead, mergeToRecordId, mergedRecordId, updatedContent, performParentingChecks);
        }

        #endregion

        #region | Private Methods |

        EntityReference GetTransactionCurrency(string entityLogicalName, Guid id)
        {
            TransactionCurrencyHelper currencyHelper = new TransactionCurrencyHelper(this.OrganizationService);
            return currencyHelper.GetTransactionCurrencyInfo(id, entityLogicalName);
        }

        #endregion
    }
}
