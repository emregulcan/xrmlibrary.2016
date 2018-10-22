using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Security
{
    /// <summary>
    /// A <c>Role (security role)</c> represents a grouping of security privileges. Users are assigned roles that authorize their access to Microsoft Dynamics CRM. 
    /// This class provides mostly used common methods for <c>Role</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334231(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class RoleHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"></param>
        public RoleHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334231(v=crm.8).aspx";
            this.EntityName = "role";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Get <c>Role</c> details.
        /// </summary>
        /// <param name="id"></param>
        /// <returns>
        /// <see cref="Entity"/> for <c>Role</c> details
        /// </returns>
        public Entity GetById(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            return this.OrganizationService.Retrieve(this.EntityName, id, new ColumnSet(true));
        }

        /// <summary>
        /// Get <c>Role</c> details.
        /// </summary>
        /// <param name="roleName"></param>
        /// <param name="retrieveColumns">
        /// Default is <c>ColumnSet(true)</c>.
        /// </param>
        /// <returns>
        /// <see cref="Entity"/> for <c>Role</c> details
        /// </returns>
        public Entity GetByName(string roleName, params string[] retrieveColumns)
        {
            ExceptionThrow.IfNullOrEmpty(roleName, "roleName");

            Entity result = null;

            QueryExpression query = new QueryExpression("role");
            query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, roleName));
            query.ColumnSet = !retrieveColumns.IsNullOrEmpty() ? new ColumnSet(retrieveColumns) : new ColumnSet(true);

            var serviceResponse = this.Get(query);

            if (serviceResponse != null && !serviceResponse.Entities.IsNullOrEmpty())
            {
                result = serviceResponse.Entities[0];
            }

            return result;
        }

        /// <summary>
        /// Get <c>Role</c> Id by specified name.
        /// </summary>
        /// <param name="roleName"></param>
        /// <returns><c>Role</c> Id (<see cref="Guid"/>)</returns>
        public Guid GetId(string roleName)
        {
            ExceptionThrow.IfNullOrEmpty(roleName, "roleName");

            Guid result = Guid.Empty;

            var serviceResponse = this.GetByName(roleName, "roleid");

            if (serviceResponse != null && !serviceResponse.Id.IsGuidEmpty())
            {
                result = serviceResponse.Id;
            }

            return result;
        }

        /// <summary>
        /// Add a existing <c>Privilege</c> to a <c>Role</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addprivilegesrolerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="roleId"></param>
        /// <param name="privilege"></param>
        /// <returns><see cref="AddPrivilegesRoleResponse"/></returns>
        public AddPrivilegesRoleResponse AddPrivilege(Guid roleId, RolePrivilege privilege)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfNull(privilege, "privilege");

            AddPrivilegesRoleRequest request = new AddPrivilegesRoleRequest()
            {
                RoleId = roleId,
                Privileges = new RolePrivilege[] { privilege }
            };

            return (AddPrivilegesRoleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add set of <c>Privilege</c> to a <c>Role</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addprivilegesrolerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="roleId"></param>
        /// <param name="privilegeList"></param>
        /// <returns>
        /// <see cref="AddPrivilegesRoleResponse"/>
        /// </returns>
        public AddPrivilegesRoleResponse AddPrivilege(Guid roleId, List<RolePrivilege> privilegeList)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfNullOrEmpty(privilegeList, "privilegeList");

            AddPrivilegesRoleRequest request = new AddPrivilegesRoleRequest()
            {
                RoleId = roleId,
                Privileges = privilegeList.ToArray()
            };

            return (AddPrivilegesRoleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add a existing <c>Privilege</c> with name (<c>prv</c> prefix) to a <c>Role</c>.
        /// </summary>
        /// <param name="roleId">Role Id</param>
        /// <param name="privilegeName">Privilege name with <c>prv</c> prefix</param>
        /// <param name="depth"><see cref="PrivilegeDepth"/></param>
        /// <returns>
        /// <see cref="AddPrivilegesRoleResponse"/>
        /// </returns>
        public AddPrivilegesRoleResponse AddPrivilege(Guid roleId, string privilegeName, PrivilegeDepth depth)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfNullOrEmpty(privilegeName, "privilegeName");

            PrivilegeHelper privilegeHelper = new PrivilegeHelper(this.OrganizationService);
            var privilegeList = privilegeHelper.GetPredefinedPrivileges();

            ExceptionThrow.IfNull(privilegeList, "privilegeList", "System pre-defined privileges not found (001)");
            ExceptionThrow.IfNullOrEmpty(privilegeList.Entities, "privilegeList.Entities", "System pre-defined privileges not found (002)");

            var existingRole = privilegeList.Entities.Where(d => d["name"].Equals(privilegeName)).SingleOrDefault();
            ExceptionThrow.IfNull(existingRole, "Privilege", string.Format("'{0}' privilege not found", privilegeName));

            RolePrivilege privilege = new RolePrivilege() { Depth = depth, PrivilegeId = existingRole.Id };

            return AddPrivilege(roleId, privilege);
        }

        /// <summary>
        /// Add a existing <c>Privilege</c> with name (<c>prv</c> prefix) to a <c>Role</c>.
        /// </summary>
        /// <param name="roleName">Role name</param>
        /// <param name="privilegeName">Privilege name with <c>prv</c> prefix</param>
        /// <param name="depth"><see cref="PrivilegeDepth"/></param>
        /// <returns><see cref="AddPrivilegesRoleResponse"/></returns>
        public AddPrivilegesRoleResponse AddPrivilege(string roleName, string privilegeName, PrivilegeDepth depth)
        {
            ExceptionThrow.IfNullOrEmpty(privilegeName, "privilegeName");
            ExceptionThrow.IfNullOrEmpty(roleName, "roleName");

            Guid id = GetId(roleName);
            ExceptionThrow.IfGuidEmpty(id, "id", string.Format("'{0}' not found", roleName));

            PrivilegeHelper privilegeHelper = new PrivilegeHelper(this.OrganizationService);
            var privilegeList = privilegeHelper.GetPredefinedPrivileges();

            ExceptionThrow.IfNull(privilegeList, "privilegeList", "System pre-defined privileges not found (001)");
            ExceptionThrow.IfNullOrEmpty(privilegeList.Entities, "privilegeList.Entities", "System pre-defined privileges not found (002)");

            var existingRole = privilegeList.Entities.Where(d => d["name"].Equals(privilegeName)).SingleOrDefault();
            ExceptionThrow.IfNull(existingRole, "Privilege", string.Format("'{0}' privilege not found", privilegeName));

            RolePrivilege privilege = new RolePrivilege() { Depth = depth, PrivilegeId = existingRole.Id };

            return AddPrivilege(id, privilege);
        }

        /// <summary>
        /// Add a existing <c>Privilege</c> to a <c>Role</c>.
        /// </summary>
        /// <param name="roleId">Role Id</param>
        /// <param name="privilegeId">
        /// <c>Privilege Id</c>.
        /// If you don't know <c>Privilege Id</c>, you can call <see cref="PrivilegeHelper.GetId(string)"/> method.
        /// </param>
        /// <param name="depth"><see cref="PrivilegeDepth"/></param>
        /// <returns>
        /// <see cref="AddPrivilegesRoleResponse"/>
        /// </returns>
        public AddPrivilegesRoleResponse AddPrivilege(Guid roleId, Guid privilegeId, PrivilegeDepth depth)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfGuidEmpty(privilegeId, "privilegeId");

            RolePrivilege privilege = new RolePrivilege() { Depth = depth, PrivilegeId = privilegeId };

            return AddPrivilege(roleId, privilege);
        }

        /// <summary>
        /// Add a existing <c>Privilege</c> with name (<c>prv</c> prefix) to a <c>Role</c>.
        /// </summary>
        /// <param name="roleName">Role name</param>
        /// <param name="privilegeId">
        /// <c>Privilege Id</c>.
        /// If you don't know <c>Privilege Id</c>, you can call <see cref="PrivilegeHelper.GetId(string)"/> method.
        /// </param>
        /// <param name="depth"><see cref="PrivilegeDepth"/></param>
        /// <returns><see cref="AddPrivilegesRoleResponse"/></returns>
        public AddPrivilegesRoleResponse AddPrivilege(string roleName, Guid privilegeId, PrivilegeDepth depth)
        {
            ExceptionThrow.IfGuidEmpty(privilegeId, "privilegeId");
            ExceptionThrow.IfNullOrEmpty(roleName, "roleName");

            Guid id = GetId(roleName);
            ExceptionThrow.IfGuidEmpty(id, "id", string.Format("'{0}' not found", roleName));

            RolePrivilege privilege = new RolePrivilege() { Depth = depth, PrivilegeId = privilegeId };

            return AddPrivilege(id, privilege);
        }

        /// <summary>
        /// Remove a <c>Privilege</c> from <c>Role</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removeprivilegerolerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="roleId"></param>
        /// <param name="privilegeId"></param>
        /// <returns>
        /// <see cref="RemovePrivilegeRoleResponse"/>
        /// </returns>
        public RemovePrivilegeRoleResponse RemovePrivilege(Guid roleId, Guid privilegeId)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfGuidEmpty(privilegeId, "privilegeId");

            RemovePrivilegeRoleRequest request = new RemovePrivilegeRoleRequest()
            {
                PrivilegeId = privilegeId,
                RoleId = roleId
            };

            return (RemovePrivilegeRoleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Replace the <c>Privilege</c> set of <c>Role</c>. 
        /// This effectively deletes all existing privileges from the role and adds the new specified privileges.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.replaceprivilegesrolerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="roleId">Role Id</param>
        /// <param name="privilegeList">Privilege List</param>
        /// <returns>
        /// <see cref="ReplacePrivilegesRoleResponse"/>
        /// </returns>
        public ReplacePrivilegesRoleResponse ReplacePrivileges(Guid roleId, List<RolePrivilege> privilegeList)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");
            ExceptionThrow.IfNullOrEmpty(privilegeList, "privilegeList");

            ReplacePrivilegesRoleRequest request = new ReplacePrivilegesRoleRequest()
            {
                RoleId = roleId,
                Privileges = privilegeList.ToArray()
            };

            return (ReplacePrivilegesRoleResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Retrieve the <c>Privilege</c> that are assigned to the specified <c>Role</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.replaceprivilegesrolerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="roleId">Role Id</param>
        /// <returns>
        /// <see cref="RolePrivilege"/> list for data
        /// </returns>
        public List<RolePrivilege> GetPrivilegesByRoleId(Guid roleId)
        {
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");

            RetrieveRolePrivilegesRoleRequest request = new RetrieveRolePrivilegesRoleRequest()
            {
                RoleId = roleId
            };

            RetrieveRolePrivilegesRoleResponse serviceResponse = (RetrieveRolePrivilegesRoleResponse)this.OrganizationService.Execute(request);
            return serviceResponse.RolePrivileges.ToList();
        }

        #endregion
    }
}
