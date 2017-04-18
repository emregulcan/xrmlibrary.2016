using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Customer
{
    /// <summary>
    /// A <c>Contact</c> represents a person with whom a business unit has a relationship.
    /// This class provides mostly used common methods for <c>Contact</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg327903(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class ContactHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public ContactHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg327903(v=crm.8).aspx";
            this.EntityName = "contact";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Merge two different <c>Contact</c> record
        /// </summary>
        /// <param name="mergeToRecordId"></param>
        /// <param name="mergedRecordId"></param>
        /// <param name="performParentingChecks"></param>
        /// <param name="updatedContent"></param>
        /// <returns>
        /// <see cref="MergeResponse"/>
        /// </returns>
        public MergeResponse Merge(Guid mergeToRecordId, Guid mergedRecordId, bool performParentingChecks = false, Entity updatedContent = null)
        {
            ExceptionThrow.IfGuidEmpty(mergeToRecordId, "mergeToRecordId");
            ExceptionThrow.IfGuidEmpty(mergedRecordId, "mergedRecordId");
            ExceptionThrow.IfNull(updatedContent, "updatedContent");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            return commonHelper.Merge(CommonHelper.MergeEntity.Contact, mergeToRecordId, mergedRecordId, updatedContent, performParentingChecks);
        }

        #endregion
    }
}
