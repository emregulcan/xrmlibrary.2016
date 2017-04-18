using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;
using XrmLibrary.EntityHelpers.Security;

namespace XrmLibrary.EntityHelpers.Administrator
{
    /// <summary>
    /// A <c>System User</c> is a person who has access to sign in to on-premises Microsoft Dynamics CRM or Microsoft Dynamics CRM Online
    /// This class provides mostly used common methods for <c>SystemUser</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309593(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class SystemUserHelper : BaseEntityHelper
    {
        #region | Bookmarks |

        /*
         * INFO : Bookmarks.SystemUser
         * Details for user & team --> https://msdn.microsoft.com/en-us/library/gg328485(v=crm.8).aspx
         * 
         */

        #endregion

        #region | Known Issues |

        /*
         * KNOWNISSUE : SystemUser.ChangeBusinessUnit
         * In CRM 2016, D365 SDK "SetBusinessSystemUserRequest" request message is deprecated (https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setbusinesssystemuserrequest.aspx).
         * Instead of this request message should use "UpdateRequest".
         * Although the "ReassignPrincipal" field is mandatory in the "SetBusinessSystemUserRequest" message, no action can be taken for this field with "UpdateRequest".
         * For this reason we use the method "ReassignObjectsSystemUserRequest" if an assignment is to be made to the records.
         * 
         */

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public SystemUserHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309593(v=crm.8).aspx";
            this.EntityName = "systemuser";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve the Id for the currently logged on user or the user under whose context the code is running.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.whoamirequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <returns><c>System User</c> Id (<see cref="Guid"/>)</returns>
        public Guid GetId()
        {
            var serviceResponse = (WhoAmIResponse)this.OrganizationService.Execute(new WhoAmIRequest());
            return serviceResponse.UserId;
        }

        /// <summary>
        /// Retrieve the <c>System User</c> Id for given <c>Domain Name</c>.
        /// </summary>
        /// <param name="domainname"><c>Active Directory name</c>(with domain)</param>
        /// <returns>
        /// <c>System User</c> Id (<see cref="Guid"/>)
        /// </returns>
        public Guid GetId(string domainname)
        {
            ExceptionThrow.IfNullOrEmpty(domainname, "domainname");

            Guid result = Guid.Empty;

            QueryExpression query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("systemuserid", "fullname");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, domainname.Trim());

            var serviceResponse = this.OrganizationService.RetrieveMultiple(query);

            if (serviceResponse != null && !serviceResponse.Entities.IsNullOrEmpty())
            {
                result = serviceResponse.Entities[0].Id;
            }

            return result;
        }

        /// <summary>
        /// Disable <c>System User</c> by given Id.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602914(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        public void Disable(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 1, -1);
        }

        /// <summary>
        /// Disable a <c>System User</c> by domainname.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602914(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="domainname"><c>Active Directory name</c>(with domain)</param>
        public void Disable(string domainname)
        {
            ExceptionThrow.IfNullOrEmpty(domainname, "domainname");

            Guid id = GetId(domainname.Trim());

            ExceptionThrow.IfGuidEmpty(id, "id", string.Format("{0} user not found", domainname));
            Disable(id);
        }

        /// <summary>
        /// Enable a <c>System User</c> by Id.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602914(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        public void Enable(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 0, -1);
        }

        /// <summary>
        /// Enable a <c>System User</c> by domainname.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602914(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="domainname"><c>Active Directory name</c>(with domain)</param>
        public void Enable(string domainname)
        {
            ExceptionThrow.IfNullOrEmpty(domainname, "domainname");

            Guid id = GetId(domainname.Trim());

            ExceptionThrow.IfGuidEmpty(id, "id", string.Format("{0} user not found", domainname));
            Enable(id);
        }

        /// <summary>
        /// Re-assign all records that are owned by a given user to another user or team.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.reassignobjectssystemuserrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        /// <param name="assignedToPrincipalType">Assigned to <c>principal</c> <see cref="PrincipalType"/></param>
        /// <param name="assignedToId">Assigned to <c>principal</c> Id</param>
        /// <returns><see cref="ReassignObjectsSystemUserResponse"/></returns>
        public ReassignObjectsSystemUserResponse ReAssignOwnedObjects(Guid id, PrincipalType assignedToPrincipalType, Guid assignedToId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(assignedToId, "assignedToId");

            ReassignObjectsSystemUserRequest request = new ReassignObjectsSystemUserRequest()
            {
                UserId = id,
                ReassignPrincipal = new EntityReference(assignedToPrincipalType.Description(), assignedToId)
            };

            return (ReassignObjectsSystemUserResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove the parent (manager) for given <c>System User</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removeparentrequest.aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        /// <returns><see cref="RemoveParentResponse"/></returns>
        public RemoveParentResponse RemoveParent(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            RemoveParentRequest request = new RemoveParentRequest()
            {
                Target = new EntityReference("systemuser", id)
            };

            return (RemoveParentResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Check <c>System User</c> 's access rights on specified record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveprincipalaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        /// <param name="targetRecordLogicalName">Target record <c>logical name</c> for which to retrieve access rights.</param>
        /// <param name="targetRecordId">Target record <c>Id</c> for which to retrieve access rights.</param>
        /// <returns>
        /// <c>True</c> if <c>System User</c> has any access right on the record, otherwise return <c>false</c>
        /// </returns>
        public bool HasAccessRight(Guid id, string targetRecordLogicalName, Guid targetRecordId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(targetRecordId, "entityId");
            ExceptionThrow.IfNullOrEmpty(targetRecordLogicalName, "entityName");

            bool result = false;

            RetrievePrincipalAccessRequest request = new RetrievePrincipalAccessRequest()
            {
                Principal = new EntityReference(this.EntityName, id),
                Target = new EntityReference(targetRecordLogicalName, targetRecordId)
            };

            RetrievePrincipalAccessResponse serviceResponse = (RetrievePrincipalAccessResponse)this.OrganizationService.Execute(request);

            if (serviceResponse.AccessRights != AccessRights.None)
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Retrieve the <c>access rights</c> of the specified user to the specified record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveprincipalaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"></param>
        /// <param name="targetRecordLogicalName"></param>
        /// <param name="targetRecordId"></param>
        /// <returns>
        /// <see cref="RetrievePrincipalAccessResponse"/>.
        /// You can get <c>Access Rights</c> from <see cref="RetrievePrincipalAccessResponse.AccessRights"/> propert.
        /// </returns>
        public RetrievePrincipalAccessResponse GetAccessRights(Guid id, string targetRecordLogicalName, Guid targetRecordId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(targetRecordId, "entityId");
            ExceptionThrow.IfNullOrEmpty(targetRecordLogicalName, "entityName");

            RetrievePrincipalAccessRequest request = new RetrievePrincipalAccessRequest()
            {
                Principal = new EntityReference(this.EntityName, id),
                Target = new EntityReference(targetRecordLogicalName, targetRecordId)
            };

            return (RetrievePrincipalAccessResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Check <c>System User</c> has specified role.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602961(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        /// <param name="rolename">Security Role name</param>
        /// <returns>
        /// <c>True</c> if <c>System User</c> has given role, otherwise return <c>false</c>
        /// </returns>
        public bool HasRole(Guid id, string rolename)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(rolename, "rolename");

            bool result = false;

            RoleHelper roleHelper = new RoleHelper(this.OrganizationService);
            Guid roleId = roleHelper.GetId(rolename);

            if (!roleId.IsGuidEmpty())
            {
                result = HasRole(id, roleId);
            }

            return result;
        }

        /// <summary>
        /// Check <c>System User</c> has specified role.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602961(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>System User</c> Id</param>
        /// <param name="roleId">Security Role Id</param>
        /// <returns>
        /// <c>True</c> if <c>System User</c> has given role, otherwise return <c>false</c>
        /// </returns>
        public bool HasRole(Guid id, Guid roleId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");

            bool result = false;

            LinkEntity linkEntity = new LinkEntity()
            {
                LinkFromEntityName = "systemuserroles",
                LinkFromAttributeName = "systemuserid",
                LinkToEntityName = "systemuser",
                LinkToAttributeName = "systemuserid",
                LinkCriteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("systemuserid", ConditionOperator.Equal,id)
                    }
                }
            };

            QueryExpression query = new QueryExpression()
            {
                EntityName = "role",
                ColumnSet = new ColumnSet("roleid"),
                LinkEntities =
                {
                    new LinkEntity()
                    {
                        LinkFromEntityName = "role",
                        LinkFromAttributeName = "roleid",
                        LinkToEntityName = "systemuserroles",
                        LinkToAttributeName = "roleid",
                        LinkEntities = {linkEntity}
                    }
                },
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("roleid", ConditionOperator.Equal, roleId)
                    }
                }
            };

            var serviceResponse = this.OrganizationService.RetrieveMultiple(query);

            if (serviceResponse != null && !serviceResponse.Entities.IsNullOrEmpty())
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Change <c>Business Unit</c> of specified <c>System User</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setbusinesssystemuserrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="userId"><c>System User</c> Id</param>
        /// <param name="businessUnitId">New <c>Business Unit</c> id</param>
        /// <param name="reassignedUserId">
        /// Target <c>SystemUser</c> to which the instances of entities previously owned by the user are to be assigned. 
        /// If you do not want to re-assign it, you can set <c>null</c> or <see cref="Guid.Empty"/>. In this case the records will remain on the current owner (specified <c>SystemUser</c> on <paramref name="userId"/>)
        /// </param>
        public void ChangeBusinessUnit(Guid userId, Guid businessUnitId, Guid? reassignedUserId = null)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");
            ExceptionThrow.IfGuidEmpty(businessUnitId, "businessUnitId");

            Entity entity = new Entity(this.EntityName);
            entity["systemuserid"] = userId;
            entity["businessunitid"] = new EntityReference("businessunit", businessUnitId);

            this.OrganizationService.Update(entity);

            /*
             * INFO : re-assign record to other principal.
             * Please check <KNOWNISSUE : SystemUser.ChangeBusinessUnit> in "Known Issues" region.
             */
            if (reassignedUserId.HasValue && !reassignedUserId.Value.IsGuidEmpty() && !reassignedUserId.Value.Equals(userId))
            {
                ReAssignOwnedObjects(userId, PrincipalType.SystemUser, reassignedUserId.Value);
            }
        }

        /// <summary>
        /// Retrieve all <c>System User</c> data by given <c>Business Unit</c> Id.
        /// </summary>
        /// <param name="businessunitId"><c>Business Unit</c> Id</param>
        /// <returns><see cref="EntityCollection"/> for <c>System User</c> data</returns>
        public EntityCollection GetByBusinessUnitId(Guid businessunitId)
        {
            ExceptionThrow.IfGuidEmpty(businessunitId, "businessunitId");

            QueryExpression query = new QueryExpression(this.EntityName);
            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessunitId);

            return this.Get(query);
        }

        /// <summary>
        /// Add <c>Role</c> to <c>System User</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/jj602982.aspx
        /// </para>
        /// </summary>
        /// <param name="userId"><c>System User</c> Id</param>
        /// <param name="roleId"><c>Role</c> Id.
        /// <para>
        /// You should provide organization related role Id.
        /// </para>
        /// </param>
        public void AddRole(Guid userId, Guid roleId)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");

            EntityReferenceCollection entityReferenceCollection = new EntityReferenceCollection();
            entityReferenceCollection.Add(new EntityReference("role", roleId));

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.Associate(userId, this.EntityName, entityReferenceCollection, "systemuserroles_association");
        }

        /// <summary>
        /// Remove <c>Role</c> from <c>System User</c>.
        /// </summary>
        /// <param name="userId"><c>System User</c> Id</param>
        /// <param name="roleId"><c>Role</c> Id.</param>
        public void RemoveRole(Guid userId, Guid roleId)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");
            ExceptionThrow.IfGuidEmpty(roleId, "roleId");

            EntityReferenceCollection entityReferenceCollection = new EntityReferenceCollection();
            entityReferenceCollection.Add(new EntityReference("role", roleId));

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.Disassociate(userId, this.EntityName, entityReferenceCollection, "systemuserroles_association");
        }

        #endregion
    }
}
