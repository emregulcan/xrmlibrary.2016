using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Security
{
    /// <summary>
    /// A <c>Privilege</c> is a permission to perform an action in Microsoft Dynamics CRM.
    /// This class provides mostly used common methods for <c>Privilege</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328425(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class PrivilegeHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public PrivilegeHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328425(v=crm.8).aspx";
            this.EntityName = "privilege";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve the set of <c>Privilege</c> data defined in the system.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveprivilegesetrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Privilege</c> data
        /// </returns>
        public EntityCollection GetPredefinedPrivileges()
        {
            RetrievePrivilegeSetRequest request = new RetrievePrivilegeSetRequest();
            RetrievePrivilegeSetResponse serviceResponse = (RetrievePrivilegeSetResponse)this.OrganizationService.Execute(request);
            return serviceResponse.EntityCollection;
        }

        /// <summary>
        /// Retrieve the set of <c>Privilege</c> for <c>System User</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveuserprivilegesrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="userId"></param>
        /// <returns>
        /// 
        /// </returns>
        public List<RolePrivilege> GetPrivilegesByUserId(Guid userId)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");

            RetrieveUserPrivilegesRequest request = new RetrieveUserPrivilegesRequest()
            {
                UserId = userId
            };

            RetrieveUserPrivilegesResponse serviceResponse = (RetrieveUserPrivilegesResponse)this.OrganizationService.Execute(request);
            return serviceResponse.RolePrivileges.ToList();
        }

        /// <summary>
        /// Retrieve <c>Privilege</c> Id
        /// </summary>
        /// <param name="name">Privilege name with <c>prv</c> prefix
        /// </param>
        /// <returns>Record Id</returns>
        public Guid GetId(string name)
        {
            ExceptionThrow.IfNullOrEmpty("name", name);

            Guid result = Guid.Empty;

            var privilegeList = GetPredefinedPrivileges();
            ExceptionThrow.IfNull(privilegeList, "privilegeList", "System pre-defined privileges not found (001)");
            ExceptionThrow.IfNullOrEmpty(privilegeList.Entities, "privilegeList.Entities", "System pre-defined privileges not found (002)");

            var existingRole = privilegeList.Entities.Where(d => d["name"].Equals(name)).SingleOrDefault();
            ExceptionThrow.IfNull(existingRole, "Privilege", string.Format("'{0}' privilege not found", name));

            result = existingRole.Id;

            return result;
        }

        #endregion
    }
}
