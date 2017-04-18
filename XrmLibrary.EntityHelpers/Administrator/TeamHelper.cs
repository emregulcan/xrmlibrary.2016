using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Administrator
{
    /// <summary>
    /// A <c>Team</c> is a group of users. Individual users are called team members. A team reports to a business unit and team members may be from different business units.
    /// This class provides mostly used common methods for <c>Team</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309360(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class TeamHelper : BaseEntityHelper
    {
        #region | Bookmarks |

        /*
         * INFO : Bookmarks.Team
         * Use access teams and owner teams to collaborate and share information --> https://msdn.microsoft.com/en-us/library/dn481569(v=crm.8).aspx
         */

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public TeamHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309360(v=crm.8).aspx";
            this.EntityName = "team";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>System User</c> into to the <c>Team</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addmembersteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <param name="systemuserId"><c>System User</c> Id</param>
        /// <returns><see cref="AddMembersTeamResponse"/></returns>
        public AddMembersTeamResponse AddMember(Guid teamId, Guid systemuserId)
        {
            ExceptionThrow.IfGuidEmpty(teamId, "teamId");
            ExceptionThrow.IfGuidEmpty(systemuserId, "systemuserId");

            List<Guid> memberList = new List<Guid>() { systemuserId };
            return AddMember(teamId, memberList);
        }

        /// <summary>
        /// Add bulk <c>System User</c> data into to the <c>Team</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addmembersteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <param name="systemuserIdList"><c>System User</c> Id list</param>
        /// <returns><see cref="AddMembersTeamResponse"/></returns>
        public AddMembersTeamResponse AddMember(Guid teamId, List<Guid> systemuserIdList)
        {
            ExceptionThrow.IfGuidEmpty(teamId, "teamId");
            ExceptionThrow.IfNullOrEmpty(systemuserIdList.ToArray(), "systemuserIdList");

            AddMembersTeamRequest request = new AddMembersTeamRequest()
            {
                TeamId = teamId,
                MemberIds = systemuserIdList.ToArray()
            };

            return (AddMembersTeamResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Removes specified <c>System User</c> from <c>Team</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removemembersteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <param name="systemuserId"><c>System User</c> Id</param>
        /// <returns>
        /// <see cref="RemoveMembersTeamResponse"/>
        /// </returns>
        public RemoveMembersTeamResponse RemoveMember(Guid teamId, Guid systemuserId)
        {
            ExceptionThrow.IfGuidEmpty(teamId, "teamId");
            ExceptionThrow.IfGuidEmpty(systemuserId, "systemuserId");

            List<Guid> memberList = new List<Guid>() { systemuserId };
            return RemoveMember(teamId, memberList);
        }

        /// <summary>
        /// Remove specified <c>System User</c> list from <c>Team</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removemembersteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <param name="systemuserIdList"><c>System User</c> Id list</param>
        /// <returns>
        /// <see cref="RemoveMembersTeamResponse"/>
        /// </returns>
        public RemoveMembersTeamResponse RemoveMember(Guid teamId, List<Guid> systemuserIdList)
        {
            ExceptionThrow.IfGuidEmpty(teamId, "teamId");
            ExceptionThrow.IfNullOrEmpty(systemuserIdList.ToArray(), "systemuserIdList");

            RemoveMembersTeamRequest request = new RemoveMembersTeamRequest()
            {
                TeamId = teamId,
                MemberIds = systemuserIdList.ToArray()
            };

            return (RemoveMembersTeamResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Convert <c>Owner Team</c> to <c>Access Team</c>.
        /// <para>
        /// Please note that you can convert an <c>Owner Team</c> to an <c>Access Team</c>, if the <c>Owner Team</c> does not have the records and security roles assigned to the team. 
        /// After you convert the <c>Owner Team</c> to the <c>Access Team</c>, it can not convert the team back to the owner team.
        /// It is a <c>one way process</c>. During conversion, all queues and mailboxes associated with the team are deleted.
        /// </para>
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.convertownerteamtoaccessteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <returns>
        /// <see cref="ConvertOwnerTeamToAccessTeamResponse"/>
        /// </returns>
        public ConvertOwnerTeamToAccessTeamResponse ConvertToAccessTeam(Guid teamId)
        {
            ExceptionThrow.IfGuidEmpty(teamId, "teamId");

            ConvertOwnerTeamToAccessTeamRequest request = new ConvertOwnerTeamToAccessTeamRequest()
            {
                TeamId = teamId
            };

            return (ConvertOwnerTeamToAccessTeamResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Change the parent <c>Business Unit</c> for specified <c>Team</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setparentteamrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="teamId"><c>Team</c> Id</param>
        /// <param name="businessUnitId"><c>Business Unit</c> Id</param>
        public void ChangeParent(Guid teamId, Guid businessUnitId)
        {
            //TODO: deprecated method -- CHECK THIS
            //TODO: NOT WORKING

            ExceptionThrow.IfGuidEmpty(teamId, "teamId");
            ExceptionThrow.IfGuidEmpty(businessUnitId, "businessUnitId");

            //Entity entity = new Entity(this.EntityName);
            //entity.Id = teamId;
            //entity["businessunitid"] = new EntityReference("businessunit", businessUnitId);

            //this.OrganizationService.Update(entity);

            SetParentTeamRequest request = new SetParentTeamRequest()
            {
                TeamId = teamId,
                BusinessId = businessUnitId
            };

            this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Check <c>Team</c> 's access rights on specified record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveprincipalaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Team</c> Id</param>
        /// <param name="targetRecordLogicalName">Target record <c>logical name</c> for which to retrieve access rights.</param>
        /// <param name="targetRecordId">Target record <c>Id</c> for which to retrieve access rights.</param>
        /// <returns>
        /// <c>True</c> if <c>Team</c> has any access right on the record, otherwise return <c>false</c>
        /// </returns>
        public bool HasAccessRight(Guid id, string targetRecordLogicalName, Guid targetRecordId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(targetRecordId, "targetRecordId");
            ExceptionThrow.IfNullOrEmpty(targetRecordLogicalName, "targetRecordLogicalName");

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
        /// Retrieve the <c>Access Rights</c> of the specified <c>Team</c> to the specified record.
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
            ExceptionThrow.IfGuidEmpty(targetRecordId, "targetRecordId");
            ExceptionThrow.IfNullOrEmpty(targetRecordLogicalName, "targetRecordLogicalName");

            RetrievePrincipalAccessRequest request = new RetrievePrincipalAccessRequest()
            {
                Principal = new EntityReference(this.EntityName, id),
                Target = new EntityReference(targetRecordLogicalName, targetRecordId)
            };

            return (RetrievePrincipalAccessResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Retrieve the <c>privileges</c> for a team.
        /// </summary>
        /// <param name="id"><c>Team</c> Id</param>
        /// <returns>
        /// <see cref="RetrieveTeamPrivilegesResponse"/>.
        /// You can get <c>Privilege</c> list with <see cref="RetrieveTeamPrivilegesResponse.RolePrivileges"/> property.
        /// </returns>
        public RetrieveTeamPrivilegesResponse GetPrivileges(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            RetrieveTeamPrivilegesRequest request = new RetrieveTeamPrivilegesRequest()
            {
                TeamId = id
            };

            return (RetrieveTeamPrivilegesResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
