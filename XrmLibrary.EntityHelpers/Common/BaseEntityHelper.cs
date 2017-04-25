using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;

namespace XrmLibrary.EntityHelpers.Common
{
    /// <summary>
    /// XrmLibrary.EntityHelpers Base class
    /// </summary>
    public class BaseEntityHelper
    {
        #region | Private Definitions |

        IOrganizationService _organizationService;

        #endregion

        #region | Properties |

        /// <summary>
        /// <see cref="IOrganizationService"/>
        /// </summary>
        public IOrganizationService OrganizationService { get { return _organizationService; } }

        /// <summary>
        /// Related entity <c>MS CRM SDK</c> url
        /// </summary>
        public string SdkHelpUrl { get; internal set; }

        /// <summary>
        /// Related entity logical name
        /// </summary>
        public string EntityName { get; internal set; }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="OrganizationService"/></param>
        public BaseEntityHelper(IOrganizationService service)
        {
            ExceptionThrow.IfNull(service, "service", "IOrganizationService can not be null");

            _organizationService = service;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve record by Id
        /// </summary>
        /// <param name="id">Record (Entity) Id</param>
        /// <param name="retrievedColumns">Set columns (attributes) that you need. If you don't set, default value is <see cref="ColumnSet(true)"/></param>
        /// <returns>
        /// <see cref="Entity"/> for record data
        /// </returns>
        public Entity Get(Guid id, params string[] retrievedColumns)
        {
            ExceptionThrow.IfNullOrEmpty(this.EntityName, "BaseEntityHelper.EntityName");
            ExceptionThrow.IfGuidEmpty(id, "id");

            var columnSet = !retrievedColumns.IsNullOrEmpty() ? new ColumnSet(retrievedColumns) : new ColumnSet(true);
            return this.OrganizationService.Retrieve(this.EntityName.ToLower().Trim(), id, columnSet);
        }

        /// <summary>
        /// Retrieve records by specified <see cref="QueryBase"/> .
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.messages.retrievemultiplerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="query"></param>
        /// <returns>
        /// <see cref="EntityCollection"/> for data
        /// </returns>
        public EntityCollection Get(QueryBase query)
        {
            ExceptionThrow.IfNullOrEmpty(this.EntityName, "BaseEntityHelper.EntityName");

            return this.OrganizationService.RetrieveMultiple(query);
        }

        /// <summary>
        /// Retrieve records by specified <c>FetchXml</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.fetchxmltoqueryexpressionrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="fetchxml"><c>FetchXml</c></param>
        /// <returns>
        /// <see cref="EntityCollection"/> for data
        /// </returns>
        public EntityCollection Get(string fetchxml)
        {
            ExceptionThrow.IfNullOrEmpty(fetchxml, "fetchxml");

            FetchExpression request = new FetchExpression(fetchxml);
            return this.OrganizationService.RetrieveMultiple(request);
        }

        /// <summary>
        /// Delete an existing record
        /// </summary>
        /// <param name="id">Record Id</param>
        public void Delete(Guid id)
        {
            ExceptionThrow.IfNullOrEmpty(this.EntityName, "BaseEntityHelper.EntityName");
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.Delete(id, this.EntityName);
        }

        #endregion
    }
}
