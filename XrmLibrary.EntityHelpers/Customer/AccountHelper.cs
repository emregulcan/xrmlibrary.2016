using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Customer
{
    /// <summary>
    /// An <c>Account</c> represents a person or business to which the salesperson sells a product or service.
    /// This class provides mostly used common methods for <c>Account</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328244(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class AccountHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public AccountHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328244(v=crm.8).aspx";
            this.EntityName = "account";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Merge two different <c>Account</c> record.
        /// </summary>
        /// <param name="mergeToRecordId"></param>
        /// <param name="mergedRecordId"></param>
        /// <param name="updatedContent"></param>
        /// <param name="performParentingChecks"></param>
        /// <returns>
        /// <see cref="MergeResponse"/>
        /// </returns>
        public MergeResponse Merge(Guid mergeToRecordId, Guid mergedRecordId, Entity updatedContent, bool performParentingChecks = false)
        {
            ExceptionThrow.IfGuidEmpty(mergeToRecordId, "mergeToRecordId");
            ExceptionThrow.IfGuidEmpty(mergedRecordId, "mergedRecordId");
            ExceptionThrow.IfNull(updatedContent, "updatedContent");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            return commonHelper.Merge(CommonHelper.MergeEntity.Account, mergeToRecordId, mergedRecordId, updatedContent, performParentingChecks);
        }

        #endregion
    }
}
