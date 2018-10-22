using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Administrator;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.BusinessManagement
{
    /// <summary>
    /// <c>Queue</c>s are instrumental in organizing, prioritizing, and monitoring the progress of your work while you are using Microsoft Dynamics 365.
    /// This class provides mostly used common methods for <c>Queue</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328459(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class QueueHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public QueueHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328459(v=crm.8).aspx";
            this.EntityName = "queue";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add the specified <c>Principal</c> to the list of <c>Queue</c> members. 
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addprincipaltoqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="queueId"><c>Queue</c> Id</param>
        /// <param name="principalType"><c>System User</c> or <c>Team</c>. If the passed-in <c>PrincipalType.Team</c> , add each team member to the queue.</param>
        /// <param name="principalId"><c>Principal</c> Id</param>
        /// <returns>
        /// <see cref="AddPrincipalToQueueResponse"/>
        /// </returns>
        public AddPrincipalToQueueResponse AddPrincipalToQueue(Guid queueId, PrincipalType principalType, Guid principalId)
        {
            ExceptionThrow.IfGuidEmpty(queueId, "queueId");
            ExceptionThrow.IfGuidEmpty(principalId, "principalId");

            Entity principalEntity = null;

            switch (principalType)
            {
                case PrincipalType.SystemUser:
                    SystemUserHelper systemuserHelper = new SystemUserHelper(this.OrganizationService);
                    principalEntity = systemuserHelper.Get(principalId, "fullname");
                    break;

                case PrincipalType.Team:
                    TeamHelper teamHelper = new TeamHelper(this.OrganizationService);
                    principalEntity = teamHelper.Get(principalId, "name");
                    break;
            }

            ExceptionThrow.IfNull(principalEntity, "principal", string.Format("Principal not found with '{0}'", principalId.ToString()));

            AddPrincipalToQueueRequest request = new AddPrincipalToQueueRequest()
            {
                QueueId = queueId,
                Principal = principalEntity
            };

            return (AddPrincipalToQueueResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Retrieve all private <c>Queue</c> data of a specified <c>System User</c> and optionally all public queues.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveuserqueuesrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="systemuserId"><c>System User</c> Id</param>
        /// <param name="includePublic">Set <c>true</c>, if you need all <c>Queue</c> data (with private and public). Otherwise set <c>false</c> </param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Queue</c> data
        /// </returns>
        public EntityCollection GetBySystemUserId(Guid systemuserId, bool includePublic = false)
        {
            ExceptionThrow.IfGuidEmpty(systemuserId, "systemuserId");

            RetrieveUserQueuesRequest request = new RetrieveUserQueuesRequest()
            {
                UserId = systemuserId,
                IncludePublic = includePublic
            };

            RetrieveUserQueuesResponse serviceResponse = (RetrieveUserQueuesResponse)this.OrganizationService.Execute(request);
            return serviceResponse.EntityCollection;
        }

        #endregion
    }
}
