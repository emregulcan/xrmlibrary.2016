using System;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Service
{
    /// <summary>
    /// A <c>Contract Detail</c> contains a detailed description of the type of service to which a customer is entitled.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328309(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class ContractDetailHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"></param>
        public ContractDetailHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328309(v=crm.8).aspx";
            this.EntityName = "contractdetail";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>Contract Detail</c> to <c>Contract</c> with required data.
        /// </summary>
        /// <param name="contractId"><c>Contract</c> Id</param>
        /// <param name="customerType"><see cref="CustomerType"/></param>
        /// <param name="customerId">Customer Id</param>
        /// <param name="title">Title (description) of contract line</param>
        /// <param name="startDate">Start date</param>
        /// <param name="endDate">End date</param>
        /// <param name="totalPrice">Total price</param>
        /// <param name="totalAllotments">Total number of <c>minutes</c> or <c>case (incident)</c> allowed for the contract line.</param>
        /// <returns>
        /// Created record Id (<see cref="System.Guid"/>)
        /// </returns>
        public Guid Add(Guid contractId, CustomerType customerType, Guid customerId, string title, DateTime startDate, DateTime endDate, decimal totalPrice, int totalAllotments)
        {
            ExceptionThrow.IfGuidEmpty(contractId, "contractId");
            ExceptionThrow.IfGuidEmpty(customerId, "customerId");
            ExceptionThrow.IfNullOrEmpty(title, "title");
            ExceptionThrow.IfEquals(startDate, "startdate", DateTime.MinValue);
            ExceptionThrow.IfEquals(endDate, "endDate", DateTime.MinValue);

            Entity entity = new Entity(this.EntityName);
            entity["contractid"] = new EntityReference("contract", contractId);
            entity["title"] = title;
            entity["customerid"] = new EntityReference(customerType.Description(), customerId);
            entity["activeon"] = startDate;
            entity["expireson"] = endDate;
            entity["price"] = new Money(totalPrice);
            entity["totalallotments"] = totalAllotments;

            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add a <c>Contract Detail</c> to <c>Contract</c> with basic data.
        /// <c>Customer</c>, <c>Start Date</c> and <c>End Date</c> properties will be copied from <c>Contract</c> data.
        /// </summary>
        /// <param name="contractId"><c>Contract</c> Id</param>
        /// <param name="title">Title (description) of contract line</param>
        /// <param name="totalPrice">Total Price</param>
        /// <param name="totalAllotments">Total number of <c>minutes</c> or <c>case (incident)</c> allowed for the contract line.</param>
        /// <returns>Created record Id (<see cref="System.Guid"/>)</returns>
        public Guid Add(Guid contractId, string title, decimal totalPrice, int totalAllotments)
        {
            ExceptionThrow.IfGuidEmpty(contractId, "contractId");
            ExceptionThrow.IfNullOrEmpty(title, "title");

            ContractHelper contractHelper = new ContractHelper(this.OrganizationService);
            var contract = contractHelper.Get(contractId, "customerid", "activeon", "expireson");

            ExceptionThrow.IfNull(contract, "contract", string.Format("'{0}' not found", contractId));
            ExceptionThrow.IfNull(contract.GetAttributeValue<EntityReference>("customerid"), "contract.CustomerId");
            ExceptionThrow.IfGuidEmpty(contract.GetAttributeValue<EntityReference>("customerid").Id, "contract.CustomerId");
            ExceptionThrow.IfEquals(contract.GetAttributeValue<DateTime>("activeon"), "contract.StartDate", DateTime.MinValue);
            ExceptionThrow.IfEquals(contract.GetAttributeValue<DateTime>("expireson"), "contract.EndDate", DateTime.MinValue);

            Entity entity = new Entity(this.EntityName);
            entity["contractid"] = new EntityReference("contract", contractId);
            entity["title"] = title;
            entity["customerid"] = contract.GetAttributeValue<EntityReference>("customerid");
            entity["activeon"] = contract.GetAttributeValue<DateTime>("activeon");
            entity["expireson"] = contract.GetAttributeValue<DateTime>("expireson");
            entity["price"] = new Money(totalPrice);
            entity["totalallotments"] = totalAllotments;

            return this.OrganizationService.Create(entity);
        }
        #endregion
    }
}
