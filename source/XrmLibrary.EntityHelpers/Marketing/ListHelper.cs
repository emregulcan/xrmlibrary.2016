using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
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
        #region | Enums |

        /// <summary>
        /// <c>List (marketing list)</c>  type
        /// </summary>
        public enum ListTypeCode
        {
            /// <summary>
            /// Static list
            /// </summary>
            [Description("static")]
            Static = 0,

            /// <summary>
            /// Dynamic list
            /// </summary>
            [Description("dynamic")]
            Dynamic = 1
        }

        #endregion

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

        /// <summary>
        /// Retrieve <c>members</c> from <c>List (marketing list)</c>.
        /// <para>
        /// Please note that if your <c>list</c> is <c>dynamic</c> it must has "Fullname" (for <c>Contact</c>, <c>Lead</c>) or "Name" (for <c>Account</c>) attribute in its query.
        /// Otherwise <see cref="ListMemberItemDetail.Name"/> will be <see cref="string.Empty"/>.
        /// </para>
        /// </summary>
        /// <param name="listId">Marketing List Id</param>
        /// <param name="itemPerPage">
        /// Record count per page. If marketling list has more than value, method works in loop.
        /// It's default value <c>5000</c> and recommended range is between <c>500</c> - <c>5000</c> for better performance.
        /// </param>
        /// <returns>
        /// <see cref="ListMemberResult"/> for data.
        /// </returns>
        public ListMemberResult GetMemberList(Guid listId, int itemPerPage = 5000)
        {
            ExceptionThrow.IfGuidEmpty(listId, "listId");
            ExceptionThrow.IfNegative(itemPerPage, "itemPerPage");
            ExceptionThrow.IfEquals(itemPerPage, "itemPerPage", 0);

            ListMemberResult result = new ListMemberResult();

            var list = this.OrganizationService.Retrieve(this.EntityName, listId, new ColumnSet("type", "membertype", "query"));

            if (list != null)
            {
                ListTypeCode listType = (ListTypeCode)Convert.ToInt32(list.GetAttributeValue<bool>("type"));
                ListMemberTypeCode membertype = (ListMemberTypeCode)list.GetAttributeValue<int>("membertype");
                string query = list.GetAttributeValue<string>("query");

                switch (listType)
                {
                    case ListTypeCode.Static:
                        result = PopulateStaticList(listId, membertype, itemPerPage);
                        break;

                    case ListTypeCode.Dynamic:
                        result = PopulateDynamicList(query, membertype, itemPerPage);
                        break;
                }
            }

            return result;
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Change marketing list lock status (locked - unlocked).
        /// </summary>
        /// <param name="id"></param>
        /// <param name="isLocked"></param>
        void ChangeLockStatus(Guid id, bool isLocked)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            Entity entity = new Entity(this.EntityName);
            entity["listid"] = id;
            entity["lockstatus"] = isLocked;

            this.OrganizationService.Update(entity);
        }

        /// <summary>
        /// Update marketing list name (<c>listname</c> attribute).
        /// </summary>
        /// <param name="id"></param>
        /// <param name="name"></param>
        void UpdateName(Guid id, string name)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(name, "name");

            Entity entity = new Entity(this.EntityName);
            entity["listid"] = id;
            entity["listname"] = name;

            this.OrganizationService.Update(entity);
        }

        /// <summary>
        /// Retrieve member list from static marketing list.
        /// </summary>
        /// <param name="listId"></param>
        /// <param name="membertype"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        ListMemberResult PopulateStaticList(Guid listId, ListMemberTypeCode membertype, int count)
        {
            ListMemberResult result = new ListMemberResult();

            string[] targetColums = membertype == ListMemberTypeCode.Account ? new[] { "name" } : new[] { "fullname" };
            string nameAttribute = membertype == ListMemberTypeCode.Account ? "members.name" : "members.fullname";

            QueryExpression query = new QueryExpression("listmember");
            query.ColumnSet = new ColumnSet("entityid");
            query.Criteria.AddCondition("listid", ConditionOperator.Equal, listId);

            LinkEntity linkEntityAccount = new LinkEntity()
            {
                LinkFromEntityName = "listmember",
                LinkFromAttributeName = "entityid",
                LinkToEntityName = membertype.Description(),
                LinkToAttributeName = string.Format("{0}id", membertype.Description()),
                JoinOperator = JoinOperator.Inner,
                Columns = new ColumnSet(targetColums),
                EntityAlias = "members"
            };

            query.LinkEntities.Add(linkEntityAccount);

            result = FetchData(query, membertype, nameAttribute, count, true);

            return result;
        }

        /// <summary>
        /// Retrieve member list from dynamic marketing list.
        /// Convert fetchxml to <see cref="QueryExpression"/>.
        /// </summary>
        /// <param name="fetxhxml"></param>
        /// <param name="membertype"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        ListMemberResult PopulateDynamicList(string fetxhxml, ListMemberTypeCode membertype, int count)
        {
            ExceptionThrow.IfNullOrEmpty(fetxhxml, "fetxhxml");

            ListMemberResult result = new ListMemberResult();

            FetchXmlToQueryExpressionRequest request = new FetchXmlToQueryExpressionRequest()
            {
                FetchXml = fetxhxml
            };

            FetchXmlToQueryExpressionResponse serviceResponse = (FetchXmlToQueryExpressionResponse)this.OrganizationService.Execute(request);

            string nameAttribute = membertype == ListMemberTypeCode.Account ? "name" : "fullname";

            result = FetchData(serviceResponse.Query, membertype, nameAttribute, count, false);

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="query"></param>
        /// <param name="membertype"></param>
        /// <param name="nameAttribute"></param>
        /// <param name="count"></param>
        /// <param name="isStatic"></param>
        /// <returns></returns>
        ListMemberResult FetchData(QueryExpression query, ListMemberTypeCode membertype, string nameAttribute, int count, bool isStatic)
        {
            ListMemberResult result = new ListMemberResult();

            string name = string.Empty;

            PagingInfo pageInfo = new PagingInfo();
            pageInfo.Count = count;
            pageInfo.PageNumber = 1;

            query.PageInfo = pageInfo;

            var data = this.OrganizationService.RetrieveMultiple(query);

            if (data != null && !data.Entities.IsNullOrEmpty())
            {
                foreach (var item in data.Entities)
                {
                    name = isStatic ? item.GetAttributeValue<AliasedValue>(nameAttribute)?.Value.ToString() : item.GetAttributeValue<string>(nameAttribute);
                    result.Add(membertype, new ListMemberItemDetail(item.Id, name));
                }
            }

            while (data.MoreRecords)
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = data.PagingCookie;
                data = this.OrganizationService.RetrieveMultiple(query);

                if (data != null && !data.Entities.IsNullOrEmpty())
                {
                    foreach (var item in data.Entities)
                    {
                        name = isStatic ? item.GetAttributeValue<AliasedValue>(nameAttribute)?.Value.ToString() : item.GetAttributeValue<string>(nameAttribute);
                        result.Add(membertype, new ListMemberItemDetail(item.Id, name));
                    }
                }
            }

            return result;
        }

        #endregion
    }
}
