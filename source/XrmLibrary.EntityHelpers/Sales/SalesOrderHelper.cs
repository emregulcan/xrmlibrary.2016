using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// A <c>SalesOrder (order)</c> is a <c>Quote</c> that has been accepted.
    /// This class provides mostly used common methods for <c>SalesOrder</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334319(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class SalesOrderHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : SalesOrder
         * 
         */

        #endregion

        #region | Private Definitions |

        bool _useUtc = false;
        string[] _salesorderColumns = new string[] { "salesorderid", "name", "ordernumber", "totalamount" };
        string[] _invoiceColumns = new string[] { "invoiceid", "name", "invoicenumber", "totalamount" };

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>SalesOrder</c> 's <c>Fulfilled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Order
        /// </para>
        /// </summary>
        public enum SalesOrderFulfilledStatusCode
        {
            CustomStatusCode = 0,
            Completed = 100001,
            Partial = 100002
        }

        /// <summary>
        /// <c>SalesOrder</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Order
        /// </para>
        /// </summary>
        public enum SalesOrderActiveStatusCode
        {
            CustomStatusCode = 0,
            New = 1,
            Pending = 2
        }

        /// <summary>
        /// <c>SalesOrder</c> 's <c>Submitted</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Order
        /// </para>
        /// </summary>
        public enum SalesOrderSubmittedStatusCode
        {
            CustomStatusCode = 0,
            InProgrress = 3
        }

        /// <summary>
        /// <c>SalesOrder</c> 's <c>Canceled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Order
        /// </para>
        /// </summary>
        public enum SalesOrderCanceledStatusCode
        {
            CustomStatusCode = 0,
            NoMoney = 4
        }

        /// <summary>
        /// <c>SalesOrder</c> 's <c>Invoiced</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Order
        /// </para>
        /// </summary>
        public enum SalesOrderInvoicedStatusCode
        {
            CustomStatusCode = 0,
            Invoiced = 10003
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        /// <param name="useUTC">
        /// Set <c>true</c>, if you want to use Utc datetime format, otherwise set <c>false</c>.
        /// This parameter's default value is <c>true</c>
        /// </param>
        public SalesOrderHelper(IOrganizationService service, bool useUTC = true) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334319(v=crm.8).aspx";
            this.EntityName = "salesorder";
            this._useUtc = useUTC;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Generate a <c>SalesOrder</c> from the specified <c>Opportunity</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.generatesalesorderfromopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="retrievedColumns">
        /// Default attributes are "salesorderid", "name", "ordernumber", "totalamount".
        /// If you need more or different attributes please set this parameter
        /// </param>
        /// <returns>
        /// Returns created <c>SalesOrder</c> in <see cref="GenerateSalesOrderFromOpportunityResponse.Entity"/> property with attributes that defined in <c>retrievedQuoteColums</c> parameter
        /// </returns>
        public GenerateSalesOrderFromOpportunityResponse CreateFromOpportunity(Guid opportunityId, params string[] retrievedColumns)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");

            string[] columns = !retrievedColumns.IsNullOrEmpty() ? retrievedColumns : _salesorderColumns;

            GenerateSalesOrderFromOpportunityRequest request = new GenerateSalesOrderFromOpportunityRequest()
            {
                OpportunityId = opportunityId,
                ColumnSet = new ColumnSet(columns)
            };

            return (GenerateSalesOrderFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Lock the <c>SalesOrder</c> pricing.
        /// Please note that when the <c>Sales Order</c> pricing is locked, changes to underlying <c>Price List (price levels)</c> do not affect the prices for a <c>Sales Order</c>
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.locksalesorderpricingrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>SalesOrder</c> Id</param>
        /// <returns>
        /// <see cref="LockSalesOrderPricingResponse"/>
        /// </returns>
        public LockSalesOrderPricingResponse LockPrice(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            LockSalesOrderPricingRequest request = new LockSalesOrderPricingRequest()
            {
                SalesOrderId = id
            };

            return (LockSalesOrderPricingResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Unlock the price per unit for the products in the specified <c>SalesOrder</c>.
        /// Please note that when the <c>SalesOrder</c> pricing is unlocked, changes to underlying <c>Price List (price levels)</c> can affect the prices for a <c>Sales Order</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.unlocksalesorderpricingrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>SalesOrder</c> Id</param>
        /// <returns>
        /// <see cref="UnlockSalesOrderPricingResponse"/>
        /// </returns>
        public UnlockSalesOrderPricingResponse UnlockPrice(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            UnlockSalesOrderPricingRequest request = new UnlockSalesOrderPricingRequest()
            {
                SalesOrderId = id
            };

            return (UnlockSalesOrderPricingResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// <c>Fullfill</c> the <c>SalesOrder</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.fulfillsalesorderrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>SalesOrder</c> Id</param>
        /// <param name="status"><see cref="SalesOrderFulfilledStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="ordercloseSubject"><c>OrderClose</c> subject</param>
        /// <param name="ordercloseDescription"><c>OrderClose</c> description</param>
        /// <param name="ordercloseActualEnd"><c>OrderClose</c> actual end date</param>
        /// <returns>
        /// <see cref="FulfillSalesOrderResponse"/>
        /// </returns>
        public FulfillSalesOrderResponse Fulfill(Guid id, SalesOrderFulfilledStatusCode status, int customStatusCode = 0, string ordercloseSubject = "", string ordercloseDescription = "", DateTime? ordercloseActualEnd = null)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == SalesOrderFulfilledStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            if (string.IsNullOrEmpty(ordercloseSubject))
            {
                var salesorderEntity = this.Get(id, "ordernumber");
                ordercloseSubject = string.Format("Salesorder Close (Fulfill) - {0}", salesorderEntity.GetAttributeValue<string>("ordernumber"));
            }

            FulfillSalesOrderRequest request = new FulfillSalesOrderRequest
            {
                OrderClose = PrepareOrderCloseEntity(id, ordercloseSubject, ordercloseDescription, ordercloseActualEnd),
                Status = new OptionSetValue(statusCode)
            };

            return (FulfillSalesOrderResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// <c>Cancel</c> the <c>SalesOrder</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.cancelsalesorderrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>SalesOrder</c> Id</param>
        /// <param name="status"><see cref="SalesOrderCanceledStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="ordercloseSubject"><c>OrderClose</c> subject</param>
        /// <param name="ordercloseDescription"><c>OrderClose</c> description</param>
        /// <param name="ordercloseActualEnd"><c>OrderClose</c> actual end date</param>
        /// <returns>
        /// <see cref="CancelSalesOrderResponse"/>
        /// </returns>
        public CancelSalesOrderResponse Cancel(Guid id, SalesOrderCanceledStatusCode status, int customStatusCode = 0, string ordercloseSubject = "", string ordercloseDescription = "", DateTime? ordercloseActualEnd = null)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == SalesOrderCanceledStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            if (string.IsNullOrEmpty(ordercloseSubject))
            {
                var salesorderEntity = this.Get(id, "ordernumber");
                ordercloseSubject = string.Format("Salesorder Close (Cancel) - {0}", salesorderEntity.GetAttributeValue<string>("ordernumber"));
            }

            CancelSalesOrderRequest request = new CancelSalesOrderRequest()
            {
                OrderClose = PrepareOrderCloseEntity(id, ordercloseSubject, ordercloseDescription, ordercloseActualEnd),
                Status = new OptionSetValue(statusCode)
            };

            return (CancelSalesOrderResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Convert a <c>SalesOrder</c> to an <c>Invoice</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.convertsalesordertoinvoicerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Salesorder</c> Id</param>
        /// <param name="retrievedColumns">
        /// Default attributes are "invoiceid", "name", "invoicenumber", "totalamount".
        /// If you need more or different attributes please set this parameter
        /// </param>
        /// <returns>
        /// Returns created <c>Invoice</c> in <see cref="ConvertSalesOrderToInvoiceResponse.Entity"/>property with attributes that defined in <c>retrievedQuoteColums</c> parameter
        /// </returns>
        public ConvertSalesOrderToInvoiceResponse ConvertToInvoice(Guid id, params string[] retrievedColumns)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            string[] columns = !retrievedColumns.IsNullOrEmpty() ? retrievedColumns : _invoiceColumns;

            ConvertSalesOrderToInvoiceRequest request = new ConvertSalesOrderToInvoiceRequest()
            {
                SalesOrderId = id,
                ColumnSet = new ColumnSet(columns)
            };

            return (ConvertSalesOrderToInvoiceResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methdods |

        Entity PrepareOrderCloseEntity(Guid id, string subject, string description, DateTime? actualEnd = null)
        {
            DateTime now = this._useUtc ? DateTime.UtcNow : DateTime.Now;

            Entity orderClose = new Entity("orderclose");
            orderClose["salesorderid"] = new EntityReference(this.EntityName, id);
            orderClose["subject"] = subject;
            orderClose["description"] = description;
            orderClose["actualend"] = actualEnd.HasValue ? actualEnd.Value : now;

            return orderClose;
        }

        #endregion
    }
}
