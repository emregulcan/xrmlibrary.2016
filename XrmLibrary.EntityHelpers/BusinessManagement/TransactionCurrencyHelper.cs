using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.BusinessManagement
{
    /// <summary>
    /// Microsoft Dynamics CRM is a multicurrency system, in which each record can be associated with its own currency. This currency is called the <c>Transaction Currency</c>.
    /// This class provides mostly used common methods for <c>TransactionCurrency</c> entities
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328355(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class TransactionCurrencyHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public TransactionCurrencyHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328355(v=crm.8).aspx";
            this.EntityName = "transactioncurrency";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve <c>Transaction Currency</c> information for specified entity.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="entityLogicalName"></param>
        /// <returns></returns>
        public EntityReference GetTransactionCurrencyInfo(Guid id, string entityLogicalName)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "entityLogicalName");

            EntityReference result = null;

            var transactionCurrency = this.OrganizationService.Retrieve(entityLogicalName, id, new ColumnSet("transactioncurrencyid"));

            if (transactionCurrency != null && transactionCurrency.Contains("transactioncurrencyid"))
            {
                result = (EntityReference)transactionCurrency["transactioncurrencyid"];
            }
            else
            {
                QueryExpression organizationQuery = new QueryExpression("organization")
                {
                    ColumnSet = new ColumnSet("basecurrencyid")
                };

                var response = this.OrganizationService.RetrieveMultiple(organizationQuery);

                if (response != null && response.Entities != null && response.Entities.Count > 0)
                {
                    result = (EntityReference)response.Entities[0]["basecurrencyid"];
                }
            }

            return result;
        }

        /// <summary>
        /// Retrieve specified <c>Transaction Currency</c> exchange rate.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveexchangeraterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="transactioncurrencyId"></param>
        /// <returns>
        /// <see cref="decimal"/> Exchange Rate
        /// </returns>
        public decimal GetExchangeRate(Guid transactioncurrencyId)
        {
            ExceptionThrow.IfGuidEmpty(transactioncurrencyId, "transactioncurrencyId");

            RetrieveExchangeRateRequest request = new RetrieveExchangeRateRequest()
            {
                TransactionCurrencyId = transactioncurrencyId
            };

            RetrieveExchangeRateResponse serviceResponse = (RetrieveExchangeRateResponse)this.OrganizationService.Execute(request);
            return serviceResponse.ExchangeRate;
        }

        #endregion
    }
}
