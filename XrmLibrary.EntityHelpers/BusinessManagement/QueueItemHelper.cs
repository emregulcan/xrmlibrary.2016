using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.BusinessManagement
{
    /// <summary>
    /// A <c>Queue Item</c> is a specific item in a queue.
    /// This class provides mostly used common methods for <c>QueueItem</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328076(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class QueueItemHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public QueueItemHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328076(v=crm.8).aspx";
            this.EntityName = "queueitem";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add item to queue.
        /// Please note that you can only add item that enabled for Queue, for more information see https://msdn.microsoft.com/en-us/library/gg328459(v=crm.8).aspx#BKMK_Enabling
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addtoqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="targetLogicalName">Queue enabled entity logical name</param>
        /// <param name="targetId">Queue enabled entity Id</param>
        /// <param name="queueId">Destination <c>Queue</c> Id</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid AddToQueue(string targetLogicalName, Guid targetId, Guid queueId)
        {
            ExceptionThrow.IfGuidEmpty(queueId, "queueId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetLogicalName, "targetLogicalName");

            AddToQueueRequest request = new AddToQueueRequest()
            {
                DestinationQueueId = queueId,
                Target = new EntityReference(targetLogicalName.ToLower(), targetId)
            };

            AddToQueueResponse serviceResponse = (AddToQueueResponse)this.OrganizationService.Execute(request);
            return serviceResponse.QueueItemId;
        }

        /// <summary>
        /// Move item (<c>QueueItem</c>) from <c>Source Queue</c> to <c>Destination Queue</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addtoqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="targetLogicalName"></param>
        /// <param name="targetId"></param>
        /// <param name="sourceQueueId"></param>
        /// <param name="destinationQueueId"></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid MoveToQueue(string targetLogicalName, Guid targetId, Guid sourceQueueId, Guid destinationQueueId)
        {
            ExceptionThrow.IfGuidEmpty(destinationQueueId, "destinationQueueId");
            ExceptionThrow.IfGuidEmpty(sourceQueueId, "sourceQueueId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetLogicalName, "targetLogicalName");

            AddToQueueRequest request = new AddToQueueRequest()
            {
                SourceQueueId = sourceQueueId,
                DestinationQueueId = destinationQueueId,
                Target = new EntityReference(targetLogicalName.ToLower(), targetId)
            };

            AddToQueueResponse serviceResponse = (AddToQueueResponse)this.OrganizationService.Execute(request);
            return serviceResponse.QueueItemId;
        }

        /// <summary>
        /// Pick item from <c>Queue</c> to <c>System User</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.pickfromqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="queueItemId"><c>Queue</c> Id</param>
        /// <param name="systemuserId"><c>System User</c> Id</param>
        /// <param name="shouldRemoved">Set <c>true</c>, if you want remove item from Queue. Otherwise set <c>false</c> </param>
        /// <returns>
        /// <see cref="PickFromQueueResponse"/>
        /// </returns>
        public PickFromQueueResponse PickFromQueue(Guid queueItemId, Guid systemuserId, bool shouldRemoved = false)
        {
            ExceptionThrow.IfGuidEmpty(queueItemId, "queueItemId");
            ExceptionThrow.IfGuidEmpty(systemuserId, "systemuserId");

            PickFromQueueRequest request = new PickFromQueueRequest()
            {
                QueueItemId = queueItemId,
                RemoveQueueItem = shouldRemoved,
                WorkerId = systemuserId
            };

            return (PickFromQueueResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Assign a <c>Queue Item</c> back to the queue owner.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.releasetoqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="queueItemId"><c>Queue Item</c> Id</param>
        /// <returns><see cref="ReleaseToQueueResponse"/></returns>
        public ReleaseToQueueResponse Release(Guid queueItemId)
        {
            ExceptionThrow.IfGuidEmpty(queueItemId, "queueItemId");

            ReleaseToQueueRequest request = new ReleaseToQueueRequest()
            {
                QueueItemId = queueItemId
            };

            return (ReleaseToQueueResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove a <c>Queue Item</c> from a <c>Queue</c>.
        /// This method is equivalent to deleting the queue item.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removefromqueuerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="queueItemId"><c>Queue Item</c> Id</param>
        /// <returns><see cref="RemoveFromQueueResponse"/></returns>
        public RemoveFromQueueResponse Remove(Guid queueItemId)
        {
            ExceptionThrow.IfGuidEmpty(queueItemId, "queueItemId");

            RemoveFromQueueRequest request = new RemoveFromQueueRequest()
            {
                QueueItemId = queueItemId
            };

            return (RemoveFromQueueResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
