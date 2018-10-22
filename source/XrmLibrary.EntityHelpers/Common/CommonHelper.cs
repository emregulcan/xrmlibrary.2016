using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using System.Linq;

namespace XrmLibrary.EntityHelpers.Common
{
    /// <summary>
    /// This includes common MS CRM SDK methods
    /// </summary>
    public class CommonHelper : BaseEntityHelper
    {
        #region | Bookmarks |

        /*
         * INFO : Bookmarks.SDK
         * Perform specialized operations using Update --> https://msdn.microsoft.com/en-us/library/dn932124(v=crm.8).aspx
         * Default status and status reason values --> https://technet.microsoft.com/en-us/library/dn531157.aspx
         */

        #endregion

        #region | Enums |

        public enum MergeEntity
        {
            [Description("account")]
            Account = 1,

            [Description("contact")]
            Contact = 2,

            [Description("lead")]
            Lead = 4,

            [Description("incident")]
            Incident = 112
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public CommonHelper(IOrganizationService service) : base(service)
        {

        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Delete an existing record
        /// </summary>
        /// <param name="id">Record Id</param>
        /// <param name="entityLogicalName">Record logical name</param>
        public void Delete(Guid id, string entityLogicalName)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");

            this.OrganizationService.Delete(entityLogicalName, id);
        }

        /// <summary>
        /// Merge the information from two entity records of the same type.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mergeEntityType">Entity type that merged</param>
        /// <param name="masterRecordId">Target (merged to) of the merge operation</param>
        /// <param name="subordinateId">Id of the entity record from which to merge data</param>
        /// <param name="updatedContent">Additional entity attributes to be set during the merge operation for accounts, contacts, or leads. This property is not applied when merging incidents</param>
        /// <param name="performParentingChecks">Indicates whether to check if the parent information is different for the two entity records.</param>
        /// <returns>
        /// <see cref="MergeResponse"/>
        /// </returns>
        public MergeResponse Merge(MergeEntity mergeEntityType, Guid masterRecordId, Guid subordinateId, Entity updatedContent, bool performParentingChecks = false)
        {
            /*
             * MergeRequest CRM SDK notları
             * https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest.aspx
             * 
             * 2015-10-01 : Entities that support MERGE operation
             *   - Lead
             *   - Account
             *   - Contact
             *   - Incident
             * 
             * Incident(Case) için kısıtlamalar
             *   - UpdateContent property kullanılamaz.
             *   https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest.updatecontent.aspx
             *   https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.mergerequest_members.aspx adresinde "MergeRequest.UpdateContent" optional görünmesine rağmen ,
             *   boş gönderilince hata veriyor. Bu nedenle "Incident" için kullanımda new Entity("incident") ataması yapıyoruz
             */

            ExceptionThrow.IfGuidEmpty(subordinateId, "mergedRecordId");
            ExceptionThrow.IfGuidEmpty(masterRecordId, "mergeToRecordId");
            ExceptionThrow.IfNull(updatedContent, "updatedContent");
            ExceptionThrow.IfNotExpectedValue(updatedContent.LogicalName, mergeEntityType.Description(), "UpdatedContent.LogicalName");

            MergeRequest request = new MergeRequest()
            {
                Target = new EntityReference(mergeEntityType.Description(), masterRecordId),
                SubordinateId = subordinateId,
                PerformParentingChecks = performParentingChecks,
                UpdateContent = updatedContent
            };

            return (MergeResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Change current state and status information for given record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setstaterequest(v=crm.8).aspx .
        /// You can also find default status and status reason values on https://technet.microsoft.com/en-us/library/dn531157.aspx
        /// </para>
        /// </summary>
        /// <param name="id">Record Id</param>
        /// <param name="entityLogicalName">Record logical name</param>
        /// <param name="stateCode">State code</param>
        /// <param name="statusCode">Status code</param>
        public void UpdateState(Guid id, string entityLogicalName, int stateCode, int statusCode)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");

            Entity entity = new Entity(entityLogicalName);
            entity.Id = id;
            entity["statecode"] = new OptionSetValue(stateCode);
            entity["statuscode"] = new OptionSetValue(statusCode);

            this.OrganizationService.Update(entity);
        }

        /// <summary>
        /// Create a link (association) between records.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.messages.associaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mainRecordId"></param>
        /// <param name="mainRecordLogicalName"></param>
        /// <param name="relatedEntities"></param>
        /// <param name="relationName"></param>
        /// <returns><see cref="AssociateResponse"/></returns>
        public AssociateResponse Associate(Guid mainRecordId, string mainRecordLogicalName, EntityReferenceCollection relatedEntities, string relationName)
        {
            ExceptionThrow.IfGuidEmpty(mainRecordId, "mainRecordId");
            ExceptionThrow.IfNullOrEmpty(mainRecordLogicalName, "mainRecordLogicalName");
            ExceptionThrow.IfNull(relatedEntities, "relatedEntities");
            ExceptionThrow.IfNullOrEmpty(relationName, "relationName");

            AssociateRequest request = new AssociateRequest()
            {
                RelatedEntities = relatedEntities,
                Target = new EntityReference(mainRecordLogicalName, mainRecordId),
                Relationship = new Relationship(relationName.Trim())
            };

            return (AssociateResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Create a link (association) between records.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.messages.associaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mainRecordId"></param>
        /// <param name="mainRecordLogicalName"></param>
        /// <param name="relatedEntityIdList"></param>
        /// <param name="relatedEntityLogicalName"></param>
        /// <param name="relationName"></param>
        /// <returns>
        /// <see cref="AssociateResponse"/>
        /// </returns>
        public AssociateResponse Associate(Guid mainRecordId, string mainRecordLogicalName, List<Guid> relatedEntityIdList, string relatedEntityLogicalName, string relationName)
        {
            ExceptionThrow.IfGuidEmpty(mainRecordId, "mainRecordId");
            ExceptionThrow.IfNullOrEmpty(mainRecordLogicalName, "mainRecordLogicalName");
            ExceptionThrow.IfNull(relatedEntityIdList, "relatedEntityIdList");
            ExceptionThrow.IfNullOrEmpty(relatedEntityLogicalName, "relatedEntityLogicalName");
            ExceptionThrow.IfNullOrEmpty(relationName, "relationName");

            EntityReferenceCollection entityCollection = new EntityReferenceCollection();

            foreach (Guid item in relatedEntityIdList)
            {
                entityCollection.Add(new EntityReference(relatedEntityLogicalName, item));
            }

            AssociateRequest request = new AssociateRequest()
            {
                RelatedEntities = entityCollection,
                Target = new EntityReference(mainRecordLogicalName, mainRecordId),
                Relationship = new Relationship(relationName.Trim())
            };

            return (AssociateResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove link (association) between records.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.messages.disassociaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mainRecordId"></param>
        /// <param name="mainRecordLogicalName"></param>
        /// <param name="relatedEntities"></param>
        /// <param name="relationName"></param>
        /// <returns>
        /// <see cref="DisassociateResponse"/>
        /// </returns>
        public DisassociateResponse Disassociate(Guid mainRecordId, string mainRecordLogicalName, EntityReferenceCollection relatedEntities, string relationName)
        {
            ExceptionThrow.IfGuidEmpty(mainRecordId, "mainRecordId");
            ExceptionThrow.IfNullOrEmpty(mainRecordLogicalName, "mainRecordLogicalName");
            ExceptionThrow.IfNull(relatedEntities, "relatedEntities");
            ExceptionThrow.IfNullOrEmpty(relationName, "relationName");

            DisassociateRequest request = new DisassociateRequest()
            {
                Target = new EntityReference(mainRecordLogicalName, mainRecordId),
                RelatedEntities = relatedEntities,
                Relationship = new Relationship(relationName)
            };

            return (DisassociateResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove link (association) between records.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.messages.disassociaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="mainRecordId"></param>
        /// <param name="mainRecordLogicalName"></param>
        /// <param name="relatedEntityIdList"></param>
        /// <param name="relatedEntityLogicalName"></param>
        /// <param name="relationName"></param>
        /// <returns>
        /// <see cref="DisassociateResponse"/>
        /// </returns>
        public DisassociateResponse Disassociate(Guid mainRecordId, string mainRecordLogicalName, List<Guid> relatedEntityIdList, string relatedEntityLogicalName, string relationName)
        {
            ExceptionThrow.IfGuidEmpty(mainRecordId, "mainRecordId");
            ExceptionThrow.IfNullOrEmpty(mainRecordLogicalName, "mainRecordLogicalName");
            ExceptionThrow.IfNull(relatedEntityIdList, "relatedEntityIdList");
            ExceptionThrow.IfNullOrEmpty(relatedEntityLogicalName, "relatedEntityLogicalName");
            ExceptionThrow.IfNullOrEmpty(relationName, "relationName");

            EntityReferenceCollection entityCollection = new EntityReferenceCollection();

            foreach (Guid item in relatedEntityIdList)
            {
                entityCollection.Add(new EntityReference(relatedEntityLogicalName, item));
            }

            DisassociateRequest request = new DisassociateRequest()
            {
                Target = new EntityReference(mainRecordLogicalName, mainRecordId),
                RelatedEntities = entityCollection,
                Relationship = new Relationship(relationName)
            };

            return (DisassociateResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Assign the specified record to a new owner (user or team).
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.assignrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="assigneeType">
        /// <see cref="PrincipalType"/>
        /// </param>
        /// <param name="assigneeId"></param>
        /// <param name="targetEntityLogicalName"></param>
        /// <param name="targetId"></param>
        public void Assign(PrincipalType assigneeType, Guid assigneeId, string targetEntityLogicalName, Guid targetId)
        {
            ExceptionThrow.IfGuidEmpty(assigneeId, "assigneeId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetEntityLogicalName, "targetEntityLogicalName");

            Entity entity = new Entity(targetEntityLogicalName.Trim());
            entity.Id = targetId;
            entity["ownerid"] = new EntityReference(assigneeType.Description(), assigneeId);

            this.OrganizationService.Update(entity);
        }

        /// <summary>
        /// Grant a security principal (user or team) access to the specified record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.grantaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="shareToPrincipal"><see cref="PrincipalType"/></param>
        /// <param name="shareToId"></param>
        /// <param name="targetEntityLogicalName"></param>
        /// <param name="targetId"></param>
        /// <param name="targetAccessRights"></param>
        /// <returns>
        /// <see cref="GrantAccessResponse"/>
        /// </returns>
        public GrantAccessResponse Share(PrincipalType shareToPrincipal, Guid shareToId, string targetEntityLogicalName, Guid targetId, AccessRights targetAccessRights)
        {
            ExceptionThrow.IfGuidEmpty(shareToId, "shareToId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetEntityLogicalName, "targetEntityLogicalName");

            GrantAccessRequest request = new GrantAccessRequest()
            {
                PrincipalAccess = new PrincipalAccess()
                {
                    Principal = new EntityReference(shareToPrincipal.Description(), shareToId),
                    AccessMask = targetAccessRights
                },
                Target = new EntityReference(targetEntityLogicalName, targetId)
            };

            return (GrantAccessResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Remove all access to a record for the specified principal.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.revokeaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="shareToPrincipal"><see cref="PrincipalType"/></param>
        /// <param name="shareToId"></param>
        /// <param name="targetEntityLogicalName"></param>
        /// <param name="targetId"></param>
        /// <returns>
        /// <see cref="RevokeAccessResponse"/>
        /// </returns>
        public RevokeAccessResponse RemoveShare(PrincipalType shareToPrincipal, Guid shareToId, string targetEntityLogicalName, Guid targetId)
        {
            ExceptionThrow.IfGuidEmpty(shareToId, "shareToId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetEntityLogicalName, "targetEntityLogicalName");

            RevokeAccessRequest request = new RevokeAccessRequest()
            {
                Revokee = new EntityReference(shareToPrincipal.Description(), shareToId),
                Target = new EntityReference(targetEntityLogicalName, targetId)
            };

            return (RevokeAccessResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Replace the access rights on the target record for the specified security principal (user or team).
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.modifyaccessrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="shareToPrincipal"><see cref="PrincipalType"/></param>
        /// <param name="shareToId"></param>
        /// <param name="targetEntityLogicalName"></param>
        /// <param name="targetId"></param>
        /// <param name="targetAccessRights"></param>
        /// <returns>
        /// <see cref="ModifyAccessResponse"/>
        /// </returns>
        public ModifyAccessResponse ModifyShare(PrincipalType shareToPrincipal, Guid shareToId, string targetEntityLogicalName, Guid targetId, AccessRights targetAccessRights)
        {
            ExceptionThrow.IfGuidEmpty(shareToId, "shareToId");
            ExceptionThrow.IfGuidEmpty(targetId, "targetId");
            ExceptionThrow.IfNullOrEmpty(targetEntityLogicalName, "targetEntityLogicalName");

            ModifyAccessRequest request = new ModifyAccessRequest()
            {
                PrincipalAccess = new PrincipalAccess()
                {
                    Principal = new EntityReference(shareToPrincipal.Description(), shareToId),
                    AccessMask = targetAccessRights
                },
                Target = new EntityReference(targetEntityLogicalName, targetId)
            };

            return (ModifyAccessResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Validates the state transition.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.isvalidstatetransitionrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="recordId">Record Id</param>
        /// <param name="entityLogicalName">Record's entity logical name</param>
        /// <param name="newStateCode"></param>
        /// <param name="newStatusCode"></param>
        /// <returns>
        /// Returns <c>true</c> if is valid. (<see cref="IsValidStateTransitionResponse.IsValid"/>)
        /// </returns>
        public bool Validate(Guid recordId, string entityLogicalName, string newStateCode, int newStatusCode)
        {
            ExceptionThrow.IfGuidEmpty(recordId, "id");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");

            List<string> supportedEntityList = new List<string>()
            {
                "incident",
                "msdyn_postalbum",
                "msdyn_postconfig",
                "msdyn_postruleconfig",
                "msdyn_wallsavedquery",
                "msdyn_wallsavedqueryusersettings",
                "opportunity"
            };

            if (!supportedEntityList.Contains(entityLogicalName.ToLower().Trim()))
            {
                ExceptionThrow.IfNotExpectedValue(entityLogicalName, "entityLogicalName", "", string.Format("{0} is not supported for this operation. For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.isvalidstatetransitionrequest(v=crm.8).aspx", entityLogicalName));
            }

            IsValidStateTransitionRequest request = new IsValidStateTransitionRequest()
            {
                Entity = new EntityReference(entityLogicalName, recordId),
                NewState = newStateCode.Trim(),
                NewStatus = newStatusCode
            };

            IsValidStateTransitionResponse response = (IsValidStateTransitionResponse)this.OrganizationService.Execute(request);
            return response.IsValid;
        }

        /// <summary>
        /// Retrieve all the entity records that are related to the specified record.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="recordId">Record Id</param>
        /// <param name="entityLogicalName">Record's entity logical name</param>
        /// <param name="query">
        /// <see cref="QueryExpression"/> query for related records.</param>
        /// <param name="rollupTypeCode">
        /// The <c>Rollup Type</c> (<see cref="RollupType"/> ) for the supported entities depends on the target entity type.
        /// <c>0 : None</c>, a rollup record is not requested. This member only retrieves the records that are directly related to a parent record.
        /// <c>1 : Related</c>, a rollup record is not requested. This member only retrieves the records that are directly related to a parent record
        /// <c>2 : Extended</c>, a rollup record that is directly related to a parent record and to any descendent record of a parent record, for example, children, grandchildren, and great-grandchildren records. 
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluptype(v=crm.8).aspx
        /// </para>
        /// </param>.
        /// <returns>
        /// <seealso cref="EntityCollection"/> for related rollup data
        /// </returns>
        public EntityCollection Rollup(Guid recordId, string entityLogicalName, QueryExpression query, RollupType rollupTypeCode)
        {
            ExceptionThrow.IfGuidEmpty(recordId, "recordId");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");
            ExceptionThrow.IfNull(query, "query");
            ExceptionThrow.IfNullOrEmpty(query.EntityName, "query.EntityName");

            string queryEntity = query.EntityName.ToLower().Trim();

            List<string> supportedTargetList = new List<string>()
            {
                "account",
                "contact",
                "opportunity"
            };

            if (!supportedTargetList.Contains(entityLogicalName.ToLower().Trim()))
            {
                ExceptionThrow.IfNotExpectedValue(entityLogicalName, "entityLogicalName", "", string.Format("'{0}' is not supported for this operation. For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx", entityLogicalName));
            }

            if (!RollupRequestSupportedEntityList().Contains(queryEntity))
            {
                ExceptionThrow.IfNotExpectedValue(queryEntity, "Query.EntityName", "", string.Format("'{0}' is not supported for this operation. For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx", queryEntity));
            }

            if (!RollupRequestSupportedRollupCombinationList().Contains(new Tuple<string, string, RollupType>(queryEntity, entityLogicalName, rollupTypeCode)))
            {
                ExceptionThrow.IfNotExpectedValue(entityLogicalName, "entityLogicalName", "", string.Format("Target entity ('{0}') - Supported Entity ('{1}') - Rollup Type ('{2}') combination is not supported for this operation. For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx", entityLogicalName, queryEntity, rollupTypeCode.ToString()));
            }

            RollupRequest request = new RollupRequest()
            {
                Target = new EntityReference(entityLogicalName, recordId),
                Query = query,
                RollupType = rollupTypeCode
            };

            RollupResponse serviceResponse = (RollupResponse)this.OrganizationService.Execute(request);
            return serviceResponse.EntityCollection;
        }

        /// <summary>
        /// Retrieve all shared principal (user or team) and <c>AccessRights</c> information (such as Read, Write etc).
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrievesharedprincipalsandaccessrequest.aspx
        /// </para>
        /// </summary>
        /// <param name="recordId">Record Id</param>
        /// <param name="entityLogicalName">Record's entity logical name</param>
        /// <returns>
        /// <seealso cref="PrincipalAccess"/> list for data
        /// </returns>
        public List<PrincipalAccess> GetSharedPrincipalsAndAccess(Guid recordId, string entityLogicalName)
        {
            ExceptionThrow.IfGuidEmpty(recordId, "recordId");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");

            List<PrincipalAccess> result = new List<PrincipalAccess>();

            RetrieveSharedPrincipalsAndAccessRequest request = new RetrieveSharedPrincipalsAndAccessRequest()
            {
                Target = new EntityReference(entityLogicalName, recordId)
            };

            var serviceResponse = (RetrieveSharedPrincipalsAndAccessResponse)this.OrganizationService.Execute(request);

            if (serviceResponse != null && serviceResponse.PrincipalAccesses != null && serviceResponse.PrincipalAccesses.Length > 0)
            {
                result = serviceResponse.PrincipalAccesses.ToList();
            }

            return result;
        }

        /// <summary>
        /// Reassign all records that are owned by the security principal (user or team) to another security principal (user or team).
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.reassignobjectsownerrequest(v=crm.7).aspx
        /// </para>
        /// </summary>
        /// <param name="fromPrincipalType"></param>
        /// <param name="fromPrincipalId"></param>
        /// <param name="toPrincipalType"></param>
        /// <param name="toPrincipalId"></param>
        /// <returns>
        /// <see cref="ReassignObjectsOwnerResponse"/> 
        /// </returns>
        public ReassignObjectsOwnerResponse ReassignOwnership(PrincipalType fromPrincipalType, Guid fromPrincipalId, PrincipalType toPrincipalType, Guid toPrincipalId)
        {
            ExceptionThrow.IfGuidEmpty(fromPrincipalId, "fromPrincipalId");
            ExceptionThrow.IfGuidEmpty(toPrincipalId, "toPrincipalId");

            ReassignObjectsOwnerRequest request = new ReassignObjectsOwnerRequest()
            {
                FromPrincipal = new EntityReference(fromPrincipalType.Description(), fromPrincipalId),
                ToPrincipal = new EntityReference(toPrincipalType.Description(), toPrincipalId)
            };

            return (ReassignObjectsOwnerResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx
        /// </summary>
        /// <returns></returns>
        List<string> RollupRequestSupportedEntityList()
        {
            return new List<string>() {
                "activitypointer",
                "annotation",
                "contract",
                "incident",
                "invoice",
                "opportunity",
                "quote",
                "salesorder"
            };
        }

        /// <summary>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.rolluprequest(v=crm.8).aspx
        /// </summary>
        /// <returns></returns>
        List<Tuple<string, string, RollupType>> RollupRequestSupportedRollupCombinationList()
        {
            return new List<Tuple<string, string, RollupType>>()
            {
                new Tuple<string, string, RollupType>("activitypointer", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("activitypointer", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("activitypointer", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("activitypointer", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("activitypointer", "opportunity", RollupType.None),
                new Tuple<string, string, RollupType>("annotation", "account", RollupType.None),
                new Tuple<string, string, RollupType>("annotation", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("annotation", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("annotation", "contact", RollupType.None),
                new Tuple<string, string, RollupType>("annotation", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("annotation", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("annotation", "opportunity", RollupType.None),
                new Tuple<string, string, RollupType>("annotation", "opportunity", RollupType.Related),
                new Tuple<string, string, RollupType>("annotation", "opportunity", RollupType.Extended),
                new Tuple<string, string, RollupType>("contract", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("contract", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("contract", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("contract", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("incident", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("incident", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("invoice", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("invoice", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("invoice", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("invoice", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("opportunity", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("opportunity", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("opportunity", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("opportunity", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("quote", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("quote", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("quote", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("quote", "contact", RollupType.Extended),
                new Tuple<string, string, RollupType>("salesorder", "account", RollupType.Related),
                new Tuple<string, string, RollupType>("salesorder", "account", RollupType.Extended),
                new Tuple<string, string, RollupType>("salesorder", "contact", RollupType.Related),
                new Tuple<string, string, RollupType>("salesorder", "contact", RollupType.Extended)
            };
        }

        #endregion
    }
}
