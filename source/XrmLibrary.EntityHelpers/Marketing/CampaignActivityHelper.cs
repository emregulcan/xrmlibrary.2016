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
    /// A <c>Campaign Activity</c> represents a step in a campaign with owner, partner, budget, timelines, and other information. <c>Campaign Activities</c> can be created for planning and running a campaign.
    /// This class provides mostly used common methods for <c>CampaignActivity</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328351(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class CampaignActivityHelper : BaseEntityHelper
    {
        #region | Enums |

        public enum ItemTypeCode
        {
            [Description("list")]
            List = 1,

            [Description("salesliterature")]
            SalesLiterature = 2
        }

        public enum CampaignActivityTypeCode
        {
            CustomTypeCode = 0,
            Research = 1,
            ContentPreparation = 2,
            TargetMarketingListCreation = 3,
            LeadQualification = 4,
            ContentDistribution = 5,
            DirectInitialContact = 6,
            DirectFollowUpContact = 7,
            ReminderDistribution = 8
        }

        public enum CampaignActivityChannelTypeCode
        {
            CustomChannelTypeCode = 0,

            [Description("phonecall")]
            Phone = 1,

            [Description("appointment")]
            Appointment = 2,

            [Description("letter")]
            Letter = 3,
            LetterViaMailMerge = 4,

            [Description("fax")]
            Fax = 5,
            FaxViaMailMerge = 6,

            [Description("email")]
            Email = 7,
            EmailWithMailMerge = 8,
            Other = 9
        }

        /// <summary>
        /// <c>CampaignActivity</c> 's <c>Open</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_CampaignActivity
        /// </para>
        /// </summary>
        public enum CampaignActivityOpenStatusCode
        {
            CustomStatusCode = -1,
            InProgress = 0,
            Proposed = 1,
            Pending = 4,
            SystemAborted = 5,
            Completed = 6
        }

        /// <summary>
        /// <c>CampaignActivity</c> 's <c>Closed</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_CampaignActivity
        /// </para>
        /// </summary>
        public enum CampaignActivityClosedStatusCode
        {
            CustomStatusCode = 0,
            Closed = 2
        }

        /// <summary>
        /// <c>CampaignActivity</c> 's <c>Canceled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_CampaignActivity
        /// </para>
        /// </summary>
        public enum CampaignActivityCanceledStatusCode
        {
            CustomStatusCode = 0,
            Canceled = 3
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public CampaignActivityHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328351(v=crm.8).aspx";
            this.EntityName = "campaignactivity";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create a <c>Campaign Activity</c>.
        /// </summary>
        /// <param name="campaignId"><c>Campaign</c> Id</param>
        /// <param name="subject">Subject</param>
        /// <param name="description">Description</param>
        /// <param name="typeCode"><see cref="CampaignActivityTypeCode"/> type code</param>
        /// <param name="channelCode"><see cref="CampaignActivityChannelTypeCode"/> channel code</param>
        /// <param name="customTypeCode">If you're using your custom typecode set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="customChannelCode">If you're using your custom channel code set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Create(Guid campaignId, string subject, string description, CampaignActivityTypeCode typeCode, CampaignActivityChannelTypeCode channelCode, int customTypeCode = 0, int customChannelCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(campaignId, "campaignId");
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            int activityTypeCode = (int)typeCode;
            int activityChannelCode = (int)channelCode;

            if (typeCode == CampaignActivityTypeCode.CustomTypeCode)
            {
                ExceptionThrow.IfNegative(customTypeCode, "customTypeCode");
                activityTypeCode = customTypeCode;
            }

            if (channelCode == CampaignActivityChannelTypeCode.CustomChannelTypeCode)
            {
                ExceptionThrow.IfNegative(customChannelCode, "customChannelCode");
                activityChannelCode = customChannelCode;
            }

            Entity entity = new Entity(this.EntityName);
            entity["regardingobjectid"] = new EntityReference("campaign", campaignId);
            entity["subject"] = subject;
            entity["typecode"] = new OptionSetValue(activityTypeCode);
            entity["channeltypecode"] = new OptionSetValue(activityChannelCode);
            entity["description"] = description;

            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add an item to a <c>Campaign Activity</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.additemcampaignactivityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignActivityId"><c>CampaignActivity</c> Id</param>
        /// <param name="itemType">
        /// <see cref="ItemTypeCode"/>
        /// This Item must be related with <c>Campaign</c> that parent of <c>Campaign Activity</c>, otherwise SDK throws an exception
        /// </param>
        /// <param name="itemId">Item Id</param>
        /// <returns>
        /// Created Item Id in <see cref="AddItemCampaignActivityResponse.CampaignActivityItemId"/> property.
        /// </returns>
        public AddItemCampaignActivityResponse AddItem(Guid campaignActivityId, ItemTypeCode itemType, Guid itemId)
        {
            ExceptionThrow.IfGuidEmpty(campaignActivityId, "campaignActivityId");
            ExceptionThrow.IfGuidEmpty(itemId, "itemId");

            AddItemCampaignActivityRequest request = new AddItemCampaignActivityRequest()
            {
                CampaignActivityId = campaignActivityId,
                ItemId = itemId,
                EntityName = itemType.Description()
            };

            return (AddItemCampaignActivityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove an item from a <c>Campaign Activity</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removeitemcampaignactivityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="campaignActivityId"><c>CampaignActivity</c> Id</param>
        /// <param name="itemId">Item Id</param>
        /// <returns><see cref="RemoveItemCampaignActivityResponse"/></returns>
        public RemoveItemCampaignActivityResponse RemoveItem(Guid campaignActivityId, Guid itemId)
        {
            ExceptionThrow.IfGuidEmpty(campaignActivityId, "campaignActivityId");
            ExceptionThrow.IfGuidEmpty(itemId, "itemId");

            RemoveItemCampaignActivityRequest request = new RemoveItemCampaignActivityRequest()
            {
                CampaignActivityId = campaignActivityId,
                ItemId = itemId
            };

            return (RemoveItemCampaignActivityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Distribute <c>CampaignActivity</c> with <c>Phonecall</c> activity.
        /// </summary>
        /// <param name="id"><c>CampaignActvity</c> Id</param>
        /// <param name="subject">Phonecall Subject</param>
        /// <param name="propagationOwnershipOptions">
        /// <c>Ownership</c> options for the activity
        /// This parameter 's values are;
        /// <c>Caller</c> : All created activities are assigned to the caller of <see cref="IOrganizationService"/>
        /// <c>ListMemberOwner</c> : Created activities are assigned to respective owners of target members.
        /// <c>None</c> : There is no change in ownership for the created activities.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.propagationownershipoptions(v=crm.8).aspx
        /// </para>
        /// </param>
        /// <param name="ownerTypeCode"></param>
        /// <param name="ownerId"></param>
        /// <param name="queueId">
        /// Set this field if you want also add created activities to an <c>Queue</c>.
        /// Method throws an exception (<see cref="ExceptionThrow.IfGuidEmpty(Guid, string, string)"/>) if you set this field to <see cref="Guid.Empty"/>
        /// </param>
        /// <param name="doesPropagate">Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.</param>
        /// <param name="useAsync">
        /// Set <c>true</c>, if you want use an asynchronous job to distribute activities. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>
        /// </param>
        /// <param name="validateActivity">
        /// Validates activity entity with required attributes and using method's enum values
        /// </param>
        /// <returns>
        /// Returns created <c>BulkOperation</c> Id in <see cref="DistributeCampaignActivityResponse.BulkOperationId"/> property.
        /// </returns>
        public DistributeCampaignActivityResponse DistributeByPhonecall(Guid id, string subject, PropagationOwnershipOptions propagationOwnershipOptions, PrincipalType ownerTypeCode, Guid ownerId, Guid? queueId, bool doesPropagate = true, bool useAsync = true, bool validateActivity = true)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            Entity entity = new Entity("phonecall");
            entity["subject"] = subject;

            return Distribute(id, entity, propagationOwnershipOptions, ownerTypeCode, ownerId, queueId, null, doesPropagate, useAsync, false, validateActivity);
        }

        /// <summary>
        /// Distribute <c>CampaignActivity</c> with <c>Fax</c> activity.
        /// </summary>
        /// <param name="id"><c>CampaignActvity</c> Id</param>
        /// <param name="subject">Fax subject</param>
        /// <param name="description">Fax description</param>
        /// <param name="propagationOwnershipOptions"></param>
        /// <param name="ownerTypeCode"></param>
        /// <param name="ownerId"></param>
        /// <param name="queueId"></param>
        /// <param name="doesPropagate">Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.</param>
        /// <param name="useAsync">
        /// Set <c>true</c>, if you want use an asynchronous job to distribute activities. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>.
        /// </param>
        /// <param name="validateActivity"></param>
        /// <returns>
        /// Returns created <c>BulkOperation</c> Id in <see cref="DistributeCampaignActivityResponse.BulkOperationId"/> property.
        /// </returns>
        public DistributeCampaignActivityResponse DistributeByFax(Guid id, string subject, string description, PropagationOwnershipOptions propagationOwnershipOptions, PrincipalType ownerTypeCode, Guid ownerId, Guid? queueId, bool doesPropagate = true, bool useAsync = true, bool validateActivity = true)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            Entity entity = new Entity("fax");
            entity["subject"] = subject;
            entity["description"] = description;

            return Distribute(id, entity, propagationOwnershipOptions, ownerTypeCode, ownerId, queueId, null, doesPropagate, useAsync, false, validateActivity);
        }

        /// <summary>
        /// Distribute <c>CampaignActivity</c> with <c>Appointment</c>.
        /// </summary>
        /// <param name="id"><c>CampaignActvity</c> Id</param>
        /// <param name="subject">Appointment Subject</param>
        /// <param name="location">Appointment Location</param>
        /// <param name="description">Appointment Description</param>
        /// <param name="scheduledStart">Appointment Start date</param>
        /// <param name="scheduledEnd">Appointment End date</param>
        /// <param name="isAllDayEvent">
        /// Set <c>true</c> if this event for all day
        /// </param>
        /// <param name="propagationOwnershipOptions"></param>
        /// <param name="ownerTypeCode"></param>
        /// <param name="ownerId"></param>
        /// <param name="queueId"></param>
        /// <param name="doesPropagate">Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.</param>
        /// <param name="useAsync">
        /// Set <c>true</c>, if you want use an asynchronous job to distribute activities. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>.
        /// </param>
        /// <param name="validateActivity"></param>
        /// <returns>
        /// Returns created <c>BulkOperation</c> Id in <see cref="DistributeCampaignActivityResponse.BulkOperationId"/> property.
        /// </returns>
        public DistributeCampaignActivityResponse DistributeByAppointment(Guid id, string subject, string location, string description, DateTime scheduledStart, DateTime scheduledEnd, bool isAllDayEvent, PropagationOwnershipOptions propagationOwnershipOptions, PrincipalType ownerTypeCode, Guid ownerId, Guid? queueId, bool doesPropagate = true, bool useAsync = true, bool validateActivity = true)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");
            ExceptionThrow.IfEquals(scheduledStart, "scheduledStart", DateTime.MinValue);
            ExceptionThrow.IfEquals(scheduledEnd, "scheduledEnd", DateTime.MinValue);

            Entity entity = new Entity("appointment");
            entity["subject"] = subject;
            entity["location"] = location;
            entity["isalldayevent"] = isAllDayEvent;
            entity["scheduledstart"] = scheduledStart;
            entity["scheduledend"] = scheduledEnd;
            entity["description"] = description;

            return Distribute(id, entity, propagationOwnershipOptions, ownerTypeCode, ownerId, queueId, null, doesPropagate, useAsync, false, validateActivity);
        }

        /// <summary>
        /// Distribute <c>CampaignActivity</c> with <c>Email</c> activity.
        /// </summary>
        /// <param name="id"><c>CampaignActvity</c> Id</param>
        /// <param name="fromType">Email <c>From</c> (<see cref="FromEntityType"/>)</param>
        /// <param name="fromId">Email <c>From</c> Id</param>
        /// <param name="subject">Email Subject</param>
        /// <param name="body">Email Body</param>
        /// <param name="emailTemplateId">
        /// If you want to use existing <c>template</c> set this field.
        /// Method throws an exception (<see cref="ExceptionThrow.IfGuidEmpty(Guid, string, string)"/>) if you set this field to <see cref="Guid.Empty"/>
        /// </param>
        /// <param name="sendEmail">
        /// Set <c>true</c> if you want to send created activities.
        /// </param>
        /// <param name="propagationOwnershipOptions"></param>
        /// <param name="ownerTypeCode"></param>
        /// <param name="ownerId"></param>
        /// <param name="queueId"></param>
        /// <param name="doesPropagate">Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.</param>
        /// <param name="useAsync">
        /// Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>, so activity will be created and executed (exp: email wil be send)
        /// </param>
        /// <param name="validateActivity"></param>
        /// <returns>
        /// Returns created <c>BulkOperation</c> Id in <see cref="DistributeCampaignActivityResponse.BulkOperationId"/> property.
        /// </returns>
        public DistributeCampaignActivityResponse DistributeByEmail(Guid id, FromEntityType fromType, Guid fromId, string subject, string body, Guid? emailTemplateId, bool sendEmail, PropagationOwnershipOptions propagationOwnershipOptions, PrincipalType ownerTypeCode, Guid ownerId, Guid? queueId, bool doesPropagate = true, bool useAsync = true, bool validateActivity = true)
        {
            ExceptionThrow.IfGuidEmpty(fromId, "fromId");

            Entity from = new Entity("activityparty");
            from["partyid"] = new EntityReference(fromType.Description(), fromId);

            Entity entity = new Entity("email");
            entity["from"] = new[] { from };

            if (emailTemplateId.HasValue)
            {
                ExceptionThrow.IfGuidEmpty(emailTemplateId.Value, "templateId");
            }
            else
            {
                ExceptionThrow.IfNullOrEmpty(subject, "subject");
                ExceptionThrow.IfNullOrEmpty(body, "body");

                entity["subject"] = subject;
                entity["description"] = body;
            }

            return Distribute(id, entity, propagationOwnershipOptions, ownerTypeCode, ownerId, queueId, emailTemplateId, doesPropagate, useAsync, sendEmail, validateActivity);
        }

        /// <summary>
        /// Distribute <c>CampaignActivity</c>.
        /// Creates the appropriate activity for each member in the list for the specified campaign activity.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.distributecampaignactivityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>CampaignActvity</c> Id</param>
        /// <param name="activityEntity"><c>Activity</c> to be distributed.
        /// This should be <c>phonecall</c> - <c>appointment</c> - <c>letter</c> - <c>fax</c> - <c>email</c>.
        /// </param>
        /// <param name="propagationOwnershipOptions">
        /// <c>Ownership</c> options for the activity
        /// This parameter 's values are;
        /// <c>Caller</c> : All created activities are assigned to the caller of <see cref="IOrganizationService"/>
        /// <c>ListMemberOwner</c> : Created activities are assigned to respective owners of target members.
        /// <c>None</c> : There is no change in ownership for the created activities.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.propagationownershipoptions(v=crm.8).aspx
        /// </para>
        /// </param>
        /// <param name="ownerTypeCode"></param>
        /// <param name="ownerId"></param>
        /// <param name="queueId">
        /// Set this field if you want also add created activities to an <c>Queue</c>.
        /// Method throws an exception (<see cref="ExceptionThrow.IfGuidEmpty(Guid, string, string)"/>) if you set this field to <see cref="Guid.Empty"/>
        /// </param>
        /// <param name="emailTemplateId">
        /// If you distribute <c>CampaignActivity</c> with <c>Email</c> and you want to use existing <c>template</c> set this field.
        /// Method throws an exception (<see cref="ExceptionThrow.IfGuidEmpty(Guid, string, string)"/>) if you set this field to <see cref="Guid.Empty"/>
        /// </param>
        /// <param name="doesPropagate">
        /// Set <c>true</c>, whether the activity is both created and executed. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>, so activity will be created and executed (exp: email wil be send)
        /// </param>
        /// <param name="useAsync">
        /// Set <c>true</c>, if you want use an asynchronous job to distribute activities. Otherwise set <c>false</c>.
        /// This field's default value is <c>true</c>
        /// </param>
        /// <param name="sendEmail">
        /// Set <c>true</c> if you distribute <c>CampaignActivity</c> with <c>Email</c> and you want to send created activities.
        /// </param>
        /// <param name="validateActivity">
        /// Validates activity entity with required attributes and using method's enum values
        /// </param>
        /// <returns>
        /// Returns created <c>BulkOperation</c> Id in <see cref="DistributeCampaignActivityResponse.BulkOperationId"/> property.
        /// </returns>
        public DistributeCampaignActivityResponse Distribute(Guid id, Entity activityEntity, PropagationOwnershipOptions propagationOwnershipOptions, PrincipalType ownerTypeCode, Guid ownerId, Guid? queueId, Guid? emailTemplateId, bool doesPropagate = true, bool useAsync = true, bool sendEmail = false, bool validateActivity = true)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(ownerId, "ownerId");

            ValidateEntity(activityEntity);

            if (validateActivity)
            {
                ValidateActivity(id, activityEntity.LogicalName);
            }

            DistributeCampaignActivityRequest request = new DistributeCampaignActivityRequest()
            {
                CampaignActivityId = id,
                Activity = activityEntity,
                Owner = new EntityReference(ownerTypeCode.Description(), ownerId),
                OwnershipOptions = propagationOwnershipOptions,
                PostWorkflowEvent = useAsync,
                Propagate = doesPropagate,
                SendEmail = sendEmail
            };

            if (queueId.HasValue)
            {
                ExceptionThrow.IfGuidEmpty(queueId.Value, "queueId");

                request.QueueId = queueId.Value;
            }

            if (emailTemplateId.HasValue)
            {
                ExceptionThrow.IfGuidEmpty(emailTemplateId.Value, "emailTemplateId");

                request.TemplateId = emailTemplateId.Value;
            }

            return (DistributeCampaignActivityResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methods |

        void ValidateActivity(Guid id, string activityName)
        {
            var entity = this.Get(id);
            var channelcode = entity.GetAttributeValue<OptionSetValue>("channeltypecode");

            ExceptionThrow.IfNull(channelcode, "channeltypecode", "CampaignActivity must has 'channeltypecode' (optionset) field.");

            if (!Enum.IsDefined(typeof(CampaignActivityChannelTypeCode), channelcode.Value))
            {
                ExceptionThrow.IfNotExpectedValue(channelcode.Value, "channelcode", 0, "'channeltypecode' value is not defined in 'CampaignActivityHelper.CampaignActivityChannelTypeCode' enum");
            }

            CampaignActivityChannelTypeCode channelEnum = (CampaignActivityChannelTypeCode)channelcode.Value;

            if (!activityName.ToLower().Trim().Equals(channelEnum.Description()))
            {
                ExceptionThrow.IfNotExpectedValue(activityName.ToLower().Trim(), "Entity.LogicalName", channelEnum.Description(), string.Format("Entity.LogicalName ('{0}') and CampaignActivity ChannelCode ('{1}') must be same value.", activityName.ToLower().Trim(), channelEnum.Description()));
            }
        }

        void ValidateEntity(Entity entity)
        {
            ExceptionThrow.IfNull(entity, "Entity");
            ExceptionThrow.IfNullOrEmpty(entity.LogicalName, "Entity.LogicalName");

            var n = entity.LogicalName.ToLower().Trim();

            if (n.Equals("phonecall") || n.Equals("appointment") || n.Equals("letter") || n.Equals("fax") || n.Equals("email"))
            {
                //Valid
            }
            else
            {
                ExceptionThrow.IfNotExpectedValue(entity.LogicalName, "Entity.LogicalName", "phonecall - appointment - letter - fax - email");
            }
        }

        #endregion
    }
}
