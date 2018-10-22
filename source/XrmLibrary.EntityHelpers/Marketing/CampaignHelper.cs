using System;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Marketing
{
    /// <summary>
    /// A <c>Campaign</c> represents a container for campaign activities and responses, sales literature, products, and lists to create, plan, execute, and track the results of a specific marketing campaign through its life. 
    /// This class provides mostly used common methods for <c>Campaign</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309509(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class CampaignHelper : BaseEntityHelper
    {
        #region | Enums |

        public enum ItemTypeCode
        {
            [Description("campaign")]
            Campaign = 1,

            [Description("list")]
            List = 2,

            [Description("product")]
            Product = 3,

            [Description("salesliterature")]
            SalesLiterature = 4
        }

        /// <summary>
        /// <c>Campaign</c> type codes
        /// </summary>
        public enum CampaignTypeCode
        {
            CustomTypeCode = 0,
            Advertisement = 1,
            DirectMarketing = 2,
            Event = 3,
            CoBranding = 4,
            Other = 5
        }

        /// <summary>
        /// <c>Campaign</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Campaign
        /// </para>
        /// </summary>
        public enum CampaignActiveStatusCode
        {
            CustomStatusCode = -1,
            Proposed = 0,
            ReadToLaunch = 1,
            Launched = 2,
            Completed = 3,
            Canceled = 4,
            Suspended = 5
        }

        /// <summary>
        /// <c>Campaign</c> 's <c>Inactive</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Campaign
        /// </para>
        /// </summary>
        public enum CampaignInctiveStatusCode
        {
            CustomStatusCode = 0,
            Inactive = 6
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public CampaignHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309509(v=crm.8).aspx";
            this.EntityName = "campaign";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create a <c>Campaign</c>.
        /// </summary>
        /// <param name="name">Name</param>
        /// <param name="code">Code.
        /// If you don't provide this data, MS CRM gives an automated number
        /// </param>
        /// <param name="typeCode"><see cref="CampaignTypeCode"/></param>
        /// <param name="description"></param>
        /// <param name="transactioncurrencyId"></param>
        /// <param name="statusCode"><see cref="CampaignActiveStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Create(string name, string code, CampaignTypeCode typeCode, string description, Guid transactioncurrencyId, CampaignActiveStatusCode statusCode, int customStatusCode = 0)
        {
            return Create(name, code, typeCode, description, transactioncurrencyId, null, null, null, null, statusCode, customStatusCode);
        }

        /// <summary>
        /// Create a <c>Campaign</c>.
        /// </summary>
        /// <param name="name">Name</param>
        /// <param name="code">Code.
        /// If you don't provide this data, MS CRM gives an automated number
        /// </param>
        /// <param name="typeCode"><see cref="CampaignTypeCode"/></param>
        /// <param name="description"></param>
        /// <param name="transactioncurrencyId"></param>
        /// <param name="actualStart">Actual Start date</param>
        /// <param name="actualEnd">Actual End date</param>
        /// <param name="statusCode"><see cref="CampaignActiveStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Create(string name, string code, CampaignTypeCode typeCode, string description, Guid transactioncurrencyId, DateTime actualStart, DateTime actualEnd, CampaignActiveStatusCode statusCode, int customStatusCode = 0)
        {
            return Create(name, code, typeCode, description, transactioncurrencyId, actualStart, actualEnd, null, null, statusCode, customStatusCode);
        }

        /// <summary>
        /// Create a <c>Campaign</c>.
        /// </summary>
        /// <param name="name">Name</param>
        /// <param name="code">Code.
        /// If you don't provide this data, MS CRM gives an automated number
        /// </param>
        /// <param name="typeCode"><see cref="CampaignTypeCode"/></param>
        /// <param name="description"></param>
        /// <param name="transactioncurrencyId"></param>
        /// <param name="actualStart">Actual Start date</param>
        /// <param name="actualEnd">Actual End date</param>
        /// <param name="proposedStart">Propesed Start date</param>
        /// <param name="proposedEnd">Proposed End date</param>
        /// <param name="status"><see cref="CampaignActiveStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Create(string name, string code, CampaignTypeCode typeCode, string description, Guid transactioncurrencyId, DateTime? actualStart, DateTime? actualEnd, DateTime? proposedStart, DateTime? proposedEnd, CampaignActiveStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfNullOrEmpty(name, "name");
            ExceptionThrow.IfGuidEmpty(transactioncurrencyId, "transactioncurrencyId");

            int statusCode = (int)status;

            if (status == CampaignActiveStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            Entity entity = new Entity(this.EntityName);
            entity["name"] = name;
            entity["typecode"] = new OptionSetValue((int)typeCode);
            entity["transactioncurrencyid"] = new EntityReference("transactioncurrency", transactioncurrencyId);
            entity["istemplate"] = false;
            entity["description"] = description;
            entity["statuscode"] = new OptionSetValue(statusCode);

            if (!string.IsNullOrEmpty(code))
            {
                entity["codename"] = code;
            }

            if (actualStart.HasValue && actualStart.Value != DateTime.MinValue)
            {
                entity["actualstart"] = actualStart.Value;
            }

            if (actualEnd.HasValue && actualEnd.Value != DateTime.MinValue)
            {
                entity["actualend"] = actualEnd.Value;
            }

            if (proposedStart.HasValue && proposedStart.Value != DateTime.MinValue)
            {
                entity["proposedstart"] = proposedStart.Value;
            }

            if (proposedEnd.HasValue && proposedEnd.Value != DateTime.MinValue)
            {
                entity["proposedend"] = proposedEnd.Value;
            }

            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add an item to <c>Campaign</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.additemcampaignrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignId"><c>Campaign</c> Id</param>
        /// <param name="itemType"><see cref="ItemTypeCode"/></param>
        /// <param name="itemId">Item Id</param>
        /// <returns>
        /// Returns created Item Id in <see cref="AddItemCampaignResponse.CampaignItemId"/> property.
        /// </returns>
        public AddItemCampaignResponse AddItem(Guid campaignId, ItemTypeCode itemType, Guid itemId)
        {
            ExceptionThrow.IfGuidEmpty(campaignId, "campaignId");
            ExceptionThrow.IfGuidEmpty(itemId, "itemId");

            AddItemCampaignRequest request = new AddItemCampaignRequest()
            {
                CampaignId = campaignId,
                EntityId = itemId,
                EntityName = itemType.Description()
            };

            return (AddItemCampaignResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove an item from <c>Campaign</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removeitemcampaignrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignId"><c>Campaign</c> Id</param>
        /// <param name="itemId">Item Id</param>
        /// <returns>
        /// <see cref="RemoveItemCampaignResponse"/>
        /// </returns>
        public RemoveItemCampaignResponse RemoveItem(Guid campaignId, Guid itemId)
        {
            ExceptionThrow.IfGuidEmpty(campaignId, "campaignId");
            ExceptionThrow.IfGuidEmpty(itemId, "itemId");

            RemoveItemCampaignRequest request = new RemoveItemCampaignRequest()
            {
                CampaignId = campaignId,
                EntityId = itemId
            };

            return (RemoveItemCampaignResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Copy <c>Campaign</c> as a new <c>Campaign</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.copycampaignrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignId"><c>Campaign</c> Id</param>
        /// <returns>
        /// Returns copied / newly created <c>Campaign</c> Id in <see cref="CopyCampaignResponse.CampaignCopyId"/> property.
        /// </returns>
        public CopyCampaignResponse CopyAsCampaign(Guid campaignId)
        {
            ExceptionThrow.IfGuidEmpty(campaignId, "campaignId");

            CopyCampaignRequest request = new CopyCampaignRequest()
            {
                BaseCampaign = campaignId,
                SaveAsTemplate = false
            };

            return (CopyCampaignResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Copy <c>Campaign</c> as a <c>Template</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.copycampaignrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignId"><c>Campaign</c> Id</param>
        /// <returns>
        /// Returns created <c>Campaign</c> Id in <see cref="CopyCampaignResponse.CampaignCopyId"/> property.
        /// </returns>
        public CopyCampaignResponse CopyAsTemplate(Guid campaignId)
        {
            ExceptionThrow.IfGuidEmpty(campaignId, "campaignId");

            CopyCampaignRequest request = new CopyCampaignRequest()
            {
                BaseCampaign = campaignId,
                SaveAsTemplate = true
            };

            return (CopyCampaignResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
