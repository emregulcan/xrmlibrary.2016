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
    /// An <c>Invoice</c> is an order that has been billed. 
    /// This class provides mostly used common methods for <c>Invoice</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg327925(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class InvoiceHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Invoice.ConvertSalesOrderToInvoiceRequest
         * SDK creates an invoice but not calculate total amount value. You should do action below to calculate total amount
         *  - Open record on CRM UI
         *  - Retrieve invoice data from SDK
         *  - Any command request (update - lock - unlock etc) from SDK
         * 
         */

        #endregion

        #region | Private Definitions |

        string[] _invoiceColumns = new string[] { "invoiceid", "name", "invoicenumber", "totalamount" };

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Invoice</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Invoice
        /// </para>
        /// </summary>
        public enum InvoiceActiveStatusCode
        {
            CustomStatusCode = 0,
            New = 1,
            PartiallyShipped = 2,
            Billed = 4,
            Booked = 5,
            Installed = 6
        }

        /// <summary>
        /// <c>Invoice</c> 's <c>Paid</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Invoice
        /// </para>
        /// </summary>
        public enum InvoicePaidStatusCode
        {
            CustomStatusCode = 0,
            Completed = 100001,
            Partial = 100002
        }

        /// <summary>
        /// <c>Invoice</c> 's <c>Cancelled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Invoice
        /// </para>
        /// </summary>
        public enum InvoiceCanceledStatusCode
        {
            CustomStatusCode = 0,
            Cancelled = 100003
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public InvoiceHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg327925(v=crm.8).aspx";
            this.EntityName = "invoice";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create an <c>Invoice</c> record from a <c>SalesOrder</c>.
        /// <para>
        /// This SDK message is originally located in <c>SalesOrder</c> entity, you can also call <seealso cref="SalesOrderHelper.ConvertToInvoice(Guid, string[])"/>
        /// </para>
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.convertsalesordertoinvoicerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <returns>
        /// Returns created <c>Invoice</c> in <see cref="ConvertSalesOrderToInvoiceResponse.Entity"/> property.
        /// </returns>
        public ConvertSalesOrderToInvoiceResponse CreateFromSalesorder(Guid salesorderId)
        {
            SalesOrderHelper salesorderHelper = new SalesOrderHelper(this.OrganizationService);
            return salesorderHelper.ConvertToInvoice(salesorderId);
        }

        /// <summary>
        /// Create an <c>Invoice</c> record from a <c>Opportunity</c> .
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.generateinvoicefromopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="retrieveColumns">
        /// Default attributes are "invoiceid", "name", "invoicenumber", "totalamount".
        /// If you need more or different attributes please set this parameter
        /// </param>
        /// <returns>
        /// Returns created <c>Invoice</c> in <see cref="GenerateInvoiceFromOpportunityResponse.Entity"/> property with attributes that defined in <c>retrievedQuoteColums</c> parameter
        /// </returns>
        public GenerateInvoiceFromOpportunityResponse CreateFromOpportunity(Guid opportunityId, params string[] retrieveColumns)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");

            string[] columns = !retrieveColumns.IsNullOrEmpty() ? retrieveColumns : _invoiceColumns;

            GenerateInvoiceFromOpportunityRequest request = new GenerateInvoiceFromOpportunityRequest()
            {
                OpportunityId = opportunityId,
                ColumnSet = new ColumnSet(columns)
            };

            return (GenerateInvoiceFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Change <c>Invoice</c> status to <c>Paid</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setstaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Invoice</c> Id</param>
        /// <param name="status"><see cref="InvoicePaidStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        public void Paid(Guid id, InvoicePaidStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == InvoicePaidStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 2, statusCode);
        }

        /// <summary>
        /// Change <c>Invoice</c> status to <c>Cancelled</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.setstaterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Invoice</c> Id</param>
        /// <param name="status"><see cref="InvoiceCanceledStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        public void Canceled(Guid id, InvoiceCanceledStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == InvoiceCanceledStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 3, statusCode);
        }

        /// <summary>
        /// Lock the <c>Price Per Unit</c> for the <c>Product</c> data in the specified <c>Invoice</c>.
        /// Please note that when the invoice pricing is locked, changes to underlying price lists (price levels) do not affect the prices for an invoice.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.lockinvoicepricingrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Invoice</c> Id</param>
        /// <returns><see cref="LockInvoicePricingResponse"/></returns>
        public LockInvoicePricingResponse LockPrice(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            LockInvoicePricingRequest request = new LockInvoicePricingRequest()
            {
                InvoiceId = id
            };

            return (LockInvoicePricingResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Unlock the <c>Price Per Unit</c> for the <c>Product</c> data in the specified <c>Invoice</c>.
        /// Please note that when the invoice pricing is unlocked, changes to underlying price lists (price level) can affect the prices for an invoice.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.unlockinvoicepricingrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Invoice</c> Id</param>
        /// <returns><see cref="UnlockInvoicePricingResponse"/></returns>
        public UnlockInvoicePricingResponse UnlockPrice(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            UnlockInvoicePricingRequest request = new UnlockInvoicePricingRequest()
            {
                InvoiceId = id
            };

            return (UnlockInvoicePricingResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
