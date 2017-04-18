using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Marketing
{
    /// <summary>
    /// <c>List (marketing list)</c> helps you create lists of potential customers or existing customers for marketing purposes. 
    /// This class provides mostly used common methods for <c>List</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309252(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class ListHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"></param>
        public ListHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309252(v=crm.8).aspx";
            this.EntityName = "list";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add members (bulk) to marketing list.
        /// Please note that you can not perform this action if marketing list is <c>locked</c>, please check <c>lockstatus</c> property.
        /// You can also call <see cref="Unlock(Guid)"/> method.
        /// </summary>
        /// <param name="listId">Marketing List Id</param>
        /// <param name="memberIdList">
        /// Member Id List.
        /// Please note that all members must be same type (<c>Account</c>, <c>Contact</c> or <c>Lead</c>), otherwise you will get an exception.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addlistmemberslistrequest(v=crm.8).aspx
        /// </para>
        /// </param>
        /// <returns><see cref="AddListMembersListResponse"/></returns>
        public AddListMembersListResponse AddMember(Guid listId, List<Guid> memberIdList)
        {
            ExceptionThrow.IfGuidEmpty(listId, "listId");
            ExceptionThrow.IfNullOrEmpty(memberIdList, "memberIdList");

            AddListMembersListRequest request = new AddListMembersListRequest()
            {
                ListId = listId,
                MemberIds = memberIdList.ToArray()
            };

            return (AddListMembersListResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add single member to marketing list.
        /// Please note that you can not perform this action if marketing list is <c>locked</c>, please check <c>lockstatus</c> property.
        /// You can also call <see cref="Unlock(Guid)"/> method.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.addmemberlistrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="listId">Marketing List Id</param>
        /// <param name="memberId">
        /// Entity Id.
        /// Please note that you can add member to list with same type (<c>Account</c>, <c>Contact</c> or <c>Lead</c>), before add member check your marketing list member type
        /// </param>
        /// <returns>
        /// Returns created memberlist Id in <see cref="AddMemberListResponse.Id"/> property.
        /// </returns>
        public AddMemberListResponse AddMember(Guid listId, Guid memberId)
        {
            ExceptionThrow.IfGuidEmpty(listId, "listId");
            ExceptionThrow.IfGuidEmpty(memberId, "memberId");

            AddMemberListRequest request = new AddMemberListRequest()
            {
                ListId = listId,
                EntityId = memberId
            };

            return (AddMemberListResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove member to marketing list.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.removememberlistrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="listId">Marketing List Id</param>
        /// <param name="memberId">Member Id</param>
        /// <returns>
        /// <see cref="RemoveMemberListResponse"/>
        /// </returns>
        public RemoveMemberListResponse RemoveMember(Guid listId, Guid memberId)
        {
            ExceptionThrow.IfGuidEmpty(listId, "listId");
            ExceptionThrow.IfGuidEmpty(memberId, "memberId");

            RemoveMemberListRequest request = new RemoveMemberListRequest()
            {
                ListId = listId,
                EntityId = memberId
            };

            return (RemoveMemberListResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Copy the members from the source list to the target list without creating duplicates. This works only with <c>Static List</c>.
        /// Please note that the type of entities contained in the source and target lists must match. If you're not sure about type of entities call <see cref="IsSameType(Guid, Guid)"/> to validate them.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.copymemberslistrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="sourceListId">Source Marketing List Id</param>
        /// <param name="targetListId">Target Marketing List Id</param>
        /// <returns>
        /// <see cref="CopyMembersListResponse"/>
        /// </returns>
        public CopyMembersListResponse CopyMembers(Guid sourceListId, Guid targetListId)
        {
            ExceptionThrow.IfGuidEmpty(sourceListId, "sourceListId");
            ExceptionThrow.IfGuidEmpty(targetListId, "targetListId");

            CopyMembersListRequest request = new CopyMembersListRequest()
            {
                SourceListId = sourceListId,
                TargetListId = targetListId
            };

            return (CopyMembersListResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Create a <c>Static List</c> from the specified dynamic list and add the members that satisfy the dynamic list query criteria to the static list.
        /// Please note that newly created static list is <c>Locked</c> and you can not add new members to created <c>Static List</c>. 
        /// If you need to unlock the list you can call <see cref="Unlock(Guid)"/> method.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.copydynamiclisttostaticrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Dynamic</c> Marketing List Id</param>
        /// <param name="createdListName">
        /// If you want set a new and known name please set this, otherwise MS CRM gives a unique name to newly created static list.
        /// </param>
        /// <returns>
        /// Returns created <c>Static</c> Marketing List Id in <see cref="CopyDynamicListToStaticResponse.StaticListId"/> property.
        /// </returns>
        public CopyDynamicListToStaticResponse CopyDynamicListToStaticList(Guid id, string createdListName)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CopyDynamicListToStaticRequest request = new CopyDynamicListToStaticRequest()
            {
                ListId = id
            };

            var result = (CopyDynamicListToStaticResponse)this.OrganizationService.Execute(request);

            if (!string.IsNullOrEmpty(createdListName))
            {
                UpdateName(result.StaticListId, createdListName);
            }

            return result;
        }

        /// <summary>
        /// <c>Lock</c> marketing list.
        /// This protects marketing list to adding new members.
        /// </summary>
        /// <param name="id">Marketing List Id</param>
        public void Lock(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            ChangeLockStatus(id, true);
        }

        /// <summary>
        /// <c>Unlock</c> marketing list.
        /// Unlocked marketing lists can be modified with add or remove members.
        /// </summary>
        /// <param name="id">Marketing List Id</param>
        public void Unlock(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            ChangeLockStatus(id, false);
        }

        /// <summary>
        /// Validate specified marketing lists with <c>Targeted At</c> field (<c>CreatedFromCode</c> attribute).
        /// </summary>
        /// <param name="sourceListId">Source Marketing List Id</param>
        /// <param name="targetListId">Target Marketing List Id</param>
        /// <returns>
        /// <c>true</c> if marketing lists are same type
        /// </returns>
        public bool IsSameType(Guid sourceListId, Guid targetListId)
        {
            ExceptionThrow.IfGuidEmpty(sourceListId, "sourceListId");
            ExceptionThrow.IfGuidEmpty(targetListId, "targetListId");

            bool result = false;

            var sourceList = this.Get(sourceListId, "createdfromcode");
            var targetList = this.Get(targetListId, "createdfromcode");

            int sourceCode = 0;
            int targetcode = 0;

            if (sourceList != null && sourceList.GetAttributeValue<OptionSetValue>("createdfromcode") != null)
            {
                sourceCode = sourceList.GetAttributeValue<OptionSetValue>("createdfromcode").Value;
            }

            if (targetList != null && targetList.GetAttributeValue<OptionSetValue>("createdfromcode") != null)
            {
                targetcode = targetList.GetAttributeValue<OptionSetValue>("createdfromcode").Value;
            }

            if (sourceCode != 0 && targetcode != 0 && sourceCode.Equals(targetcode))
            {
                result = true;
            }

            return result;
        }

        #endregion

        #region | Private Methods |

        void ChangeLockStatus(Guid id, bool isLocked)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            Entity entity = new Entity(this.EntityName);
            entity["listid"] = id;
            entity["lockstatus"] = isLocked;

            this.OrganizationService.Update(entity);
        }

        void UpdateName(Guid id, string name)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(name, "name");

            Entity entity = new Entity(this.EntityName);
            entity["listid"] = id;
            entity["listname"] = name;

            this.OrganizationService.Update(entity);
        }

        #endregion
    }
}
