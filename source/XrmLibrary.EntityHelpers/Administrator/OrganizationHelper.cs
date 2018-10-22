using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Administrator
{
    /// <summary>
    /// The <c>Organization</c> is the top level of the  Microsoft Dynamics 365 business hierarchy. 
    /// The organization can be a specific business, a holding company, a corporation, and so on. An organization is divided into business units. 
    /// This class provides mostly used common methods for <c>Organization</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309569(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class OrganizationHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public OrganizationHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309569(v=crm.8).aspx";
            this.EntityName = "organization";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve <see cref="IOrganizationService"/> <c>Caller</c> 's informations.
        /// </summary>
        /// <returns><see cref="WhoAmIResponse"/></returns>
        public WhoAmIResponse WhoAmI()
        {
            return (WhoAmIResponse)this.OrganizationService.Execute(new WhoAmIRequest());
        }

        /// <summary>
        /// Retrieve <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Business Unit</c> Id.
        /// </summary>
        /// <returns><c>Business Unit</c> Id (<see cref="Guid"/>)</returns>
        public Guid GetBusinessUnitId()
        {
            var serviceResponse = (WhoAmIResponse)this.OrganizationService.Execute(new WhoAmIRequest());
            return serviceResponse.BusinessUnitId;
        }

        /// <summary>
        /// Retrieve <see cref="IOrganizationService"/> <c>Caller</c> 's <c>Organization</c> Id.
        /// </summary>
        /// <returns><c>Organization</c> Id (<see cref="Guid"/>)</returns>
        public Guid GetOrganizationId()
        {
            var serviceResponse = (WhoAmIResponse)this.OrganizationService.Execute(new WhoAmIRequest());
            return serviceResponse.OrganizationId;
        }

        #endregion
    }
}
