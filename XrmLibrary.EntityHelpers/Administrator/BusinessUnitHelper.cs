using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Administrator
{
    /// <summary>
    /// A <c>Business Unit</c> is a unit of the top-level organization. <c>Business Unit</c>s can be parents of other business units (child business units). 
    /// This class provides mostly used common methods for <c>BusinessUnit</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309617(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class BusinessUnitHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public BusinessUnitHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309617(v=crm.8).aspx";
            this.EntityName = "businessunit";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Disable <c>Business Unit</c>.
        /// </summary>
        /// <param name="id"><c>Business Unit</c> Id</param>
        public void Disable(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 1, 2);
        }

        /// <summary>
        /// Enable <c>Business Unit</c>.
        /// </summary>
        /// <param name="id"><c>Business Unit</c> Id</param>
        public void Enable(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 0, 1);
        }

        /// <summary>
        /// Retrieve all the <c>Business Unit</c> list in the business unit hierarchy.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrievebusinesshierarchybusinessunitrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id">Parent <c>Business Unit</c> Id</param>
        /// <returns><see cref="EntityCollection"/> for <c>Business Unit</c> list</returns>
        public EntityCollection GetHierarchy(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            RetrieveBusinessHierarchyBusinessUnitRequest request = new RetrieveBusinessHierarchyBusinessUnitRequest()
            {
                ColumnSet = new ColumnSet(true),
                EntityId = id
            };

            RetrieveBusinessHierarchyBusinessUnitResponse serviceResponse = (RetrieveBusinessHierarchyBusinessUnitResponse)this.OrganizationService.Execute(request);
            return serviceResponse.EntityCollection;
        }

        /// <summary>
        /// Change the <c>Parent Business Unit</c> for specified <c>Business Unit</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setparentbusinessunitrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Business Unit</c> Id</param>
        /// <param name="parentBusinessUnitId">Parent <c>Business Unit</c> Id</param>
        public void SetParent(Guid id, Guid parentBusinessUnitId)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfGuidEmpty(parentBusinessUnitId, "parentBusinessUnitId");

            Entity entity = new Entity(this.EntityName);
            entity["businessunitid"] = id;
            entity["parentbusinessunitid"] = new EntityReference(this.EntityName, parentBusinessUnitId);

            this.OrganizationService.Update(entity);
        }

        #endregion
    }
}
