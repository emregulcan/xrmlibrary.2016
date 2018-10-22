using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Activity
{
    /// <summary>
    /// The <c>Email</c> activity lets you track and manage email communications with customers.
    /// This class provides mostly used common methods for <c>Email</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334229(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class EmailHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Email.StatusCode
         * In Microsoft Dynamics 365 (online & on-premises), the Email.StatusCode attribute cannot be null.
         * https://msdn.microsoft.com/en-us/library/gg334229(v=crm.8).aspx
         * 
         */

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public EmailHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334229(v=crm.8).aspx";
            this.EntityName = "email";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create an <c>Email</c> activity from native <c>Email</c> <see cref="Entity"/>.
        /// <para>
        /// Please note that if you want send email record please call <see cref="EmailHelper.Send(Guid, string, bool)"/> after this method.
        /// </para>
        /// </summary>
        /// <param name="entity"><see cref="Entity"/> for <c>Email Activity</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Create(Entity entity)
        {
            ValidateEmailEntity(entity);

            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Create an <c>Email</c> activity from <see cref="XrmEmail"/> object.
        /// <para>
        /// Please note that if you want send email record please call <see cref="EmailHelper.Send(Guid, string, bool)"/> after this method.
        /// </para>
        /// </summary>
        /// <param name="email"><see cref="XrmEmail"/> object</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Create(XrmEmail email)
        {
            ExceptionThrow.IfNull(email, "email");

            return Create(email.ToEntity());
        }

        /// <summary>
        /// Retrieve a <c>Tracking Token</c> that can be passed in as a parameter to the <see cref="SendEmailRequest"/> message, and also <see cref="Send(Entity, string, EntityReference)"/> method.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.gettrackingtokenemailrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <returns><c>Tracking token</c> for email 's subject (<see cref="string"/>)</returns>
        public string GetTrackingToken(string subject)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            GetTrackingTokenEmailRequest request = new GetTrackingTokenEmailRequest()
            {
                Subject = subject
            };

            GetTrackingTokenEmailResponse serviceResponse = (GetTrackingTokenEmailResponse)this.OrganizationService.Execute(request);
            return serviceResponse.TrackingToken;
        }

        /// <summary>
        /// Send or just create as 'sent' an email activity.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.sendemailrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Email Activity</c> Id</param>
        /// <param name="trackingToken">Use <see cref="GetTrackingToken(string)"/> to get <c>Tracking Token</c></param>
        /// <param name="shouldSend">
        /// Set <c>true</c>, if the email should be sent. 
        /// Otherwise set <c>false</c> to'just record as 'sent'. 
        /// This parameter's default value is <c>true</c>, so your email will be send</param>
        /// <returns>
        /// <see cref="SendEmailResponse"/>.
        /// You can retrieve <c>Subject (with Tracking Token)</c> from <see cref="SendEmailResponse.Subject"/> property.
        /// </returns>
        public SendEmailResponse Send(Guid id, string trackingToken, bool shouldSend = true)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            SendEmailRequest request = new SendEmailRequest
            {
                EmailId = id,
                IssueSend = shouldSend,
                TrackingToken = trackingToken
            };

            return (SendEmailResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Send an <c>Email</c> message using a <c>Template</c>.
        /// If you do not know your template id, you can use this method with just template title.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.sendemailfromtemplaterequest(v=crm.8).aspx 
        /// </para>
        /// </summary>
        /// <param name="email">CRM (XRM) email entity without body and subject. Email body and subject will be replaced by specified template's properties.</param>
        /// <param name="templateTitle"><c>Template</c> title</param>
        /// <param name="regarding">Regarding object that related by template</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Send(Entity email, string templateTitle, EntityReference regarding)
        {
            ExceptionThrow.IfNullOrEmpty(templateTitle, "templateTitle");
            ExceptionThrow.IfNull(regarding, "regarding");
            ExceptionThrow.IfGuidEmpty(regarding.Id, "EntityReference.Id");
            ExceptionThrow.IfNullOrEmpty(regarding.LogicalName, "EntityReference.LogicalName");

            ValidateEmailEntity(email);

            QueryExpression query = new QueryExpression("template")
            {
                ColumnSet = new ColumnSet("templateid", "templatetypecode"),
                Criteria = new FilterExpression()
            };

            query.Criteria.AddCondition("title", ConditionOperator.Equal, templateTitle);
            query.Criteria.AddCondition("templatetypecode", ConditionOperator.Equal, regarding.LogicalName.ToLower());

            EntityCollection templateCollection = this.OrganizationService.RetrieveMultiple(query);

            ExceptionThrow.IfNull(templateCollection, "template", string.Format("'{0}' template not exists or not related with '{1}' entity.", templateTitle, regarding.LogicalName));
            ExceptionThrow.IfGreaterThan(templateCollection.Entities.Count, "template", 1, "There are more than one template with same name");

            Guid templateId = templateCollection.Entities[0].Id;
            return Send(email, templateId, regarding);
        }

        /// <summary>
        /// Send an <c>Email</c> message using a <c>Template</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.sendemailfromtemplaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="email">CRM (XRM) email entity without body and subject. Email body and subject will be replaced by specified template's properties.</param>
        /// <param name="templateId"><c>Template</c> Id</param>
        /// <param name="regarding">Regarding object that related by template</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Send(Entity email, Guid templateId, EntityReference regarding)
        {
            ExceptionThrow.IfGuidEmpty(templateId, "templateId");
            ExceptionThrow.IfNull(regarding, "regarding");
            ExceptionThrow.IfGuidEmpty(regarding.Id, "EntityReference.Id");
            ExceptionThrow.IfNullOrEmpty(regarding.LogicalName, "EntityReference.LogicalName");

            ValidateEmailEntity(email);

            SendEmailFromTemplateRequest request = new SendEmailFromTemplateRequest()
            {
                Target = email,
                TemplateId = templateId,
                RegardingId = regarding.Id,
                RegardingType = regarding.LogicalName.ToLower()
            };

            var serviceResponse = (SendEmailFromTemplateResponse)this.OrganizationService.Execute(request);
            return serviceResponse.Id;
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Validate <c>Email</c> entity with required attributes.
        /// </summary>
        /// <param name="entity"></param>
        void ValidateEmailEntity(Entity entity)
        {
            ExceptionThrow.IfNull(entity, "entity");
            ExceptionThrow.IfNullOrEmpty(entity.LogicalName, "Entity.LogicalName");
            ExceptionThrow.IfNotExpectedValue(entity.LogicalName.Trim().ToLower(), "Entity.LogicalName", "email", "Entity.LogicalName must be 'email'");

            var sender = GetFrom(entity);
            var toList = GetRecipientsList(entity, "to");

            ExceptionThrow.IfNull(sender, "email.from");
            ExceptionThrow.IfGuidEmpty(sender.Id, "email.from.id");
            ExceptionThrow.IfNullOrEmpty(sender.LogicalName, "email.from.LogicalName");

            ExceptionThrow.IfNullOrEmpty(toList, "email.to");
        }

        /// <summary>
        /// Remove attribute if exists
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="node"></param>
        void RemoveAttribute(ref Entity entity, string node)
        {
            ExceptionThrow.IfNull(entity, "entity");
            ExceptionThrow.IfNullOrEmpty(node, "node");

            if (entity.Contains(node))
            {
                entity.Attributes.Remove(node);
            }
        }

        /// <summary>
        /// Retrieve <c>From</c> activityparty
        /// </summary>
        /// <param name="entity">Email entity</param>
        /// <returns></returns>
        EntityReference GetFrom(Entity entity)
        {
            EntityReference result = null;

            if (entity.Contains("from") && entity["from"] is Entity[] && ((Entity[])entity["from"]).Length > 0)
            {
                var f = ((Entity[])entity["from"])[0];

                if (f != null && f.Contains("partyid") && f["partyid"] is EntityReference)
                {
                    result = (EntityReference)f["partyid"];
                }
            }

            return result;
        }

        /// <summary>
        /// Retrieve recipient from related <c>activityparty</c> (node)
        /// </summary>
        /// <param name="entity">Email entity</param>
        /// <param name="node">"TO", "CC", "BCC"</param>
        /// <returns></returns>
        EntityReference GetRecipient(Entity entity, string node)
        {
            EntityReference result = null;
            EntityCollection nodeList = null;

            if (entity.Contains(node) && entity[node] is EntityCollection)
            {
                nodeList = (EntityCollection)entity[node];

                if (nodeList != null
                    && !nodeList.Entities.IsNullOrEmpty()
                    && nodeList.Entities[0].Contains("partyid")
                    && nodeList.Entities[0]["partyid"] is EntityReference
                    )
                {
                    result = (EntityReference)nodeList.Entities[0]["partyid"];
                }
            }

            return result;
        }

        /// <summary>
        /// Retrieve recipients list from related <c>activityparty</c> (node)
        /// </summary>
        /// <param name="entity">Email entity</param>
        /// <param name="node">"TO", "CC", "BCC"</param>
        /// <returns></returns>
        List<EntityReference> GetRecipientsList(Entity entity, string node)
        {
            List<EntityReference> result = new List<EntityReference>();

            if (entity.Contains(node) && entity[node] is EntityCollection)
            {
                var nodeList = (EntityCollection)entity[node];

                if (nodeList != null && !nodeList.Entities.IsNullOrEmpty())
                {
                    foreach (var item in nodeList.Entities)
                    {
                        if (item.Contains("partyid") && item["partyid"] is EntityReference)
                        {
                            result.Add((EntityReference)item["partyid"]);
                        }
                    }
                }
            }

            return result;
        }

        #endregion
    }
}
