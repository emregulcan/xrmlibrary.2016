using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// An <c>Invoice Product (detail)</c> stores a line item in an <c>Invoice</c>. Each invoice can have multiple invoice detail records.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334619(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class InvoiceProductHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public InvoiceProductHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334619(v=crm.8).aspx";
            this.EntityName = "invoicedetail";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> from an <c>Opportunity</c>.
        /// This retrieves the products from an <c>Opportunity</c> and copies them to the <c>Invoice</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.getinvoiceproductsfromopportunityrequest(v=crm.7).aspx
        /// </para>
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <returns>
        /// <see cref="GetInvoiceProductsFromOpportunityResponse"/>
        /// </returns>
        public GetInvoiceProductsFromOpportunityResponse Add(Guid invoiceId, Guid opportunityId)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");
            ExceptionThrow.IfGuidEmpty(invoiceId, "invoiceId");

            GetInvoiceProductsFromOpportunityRequest request = new GetInvoiceProductsFromOpportunityRequest()
            {
                OpportunityId = opportunityId,
                InvoiceId = invoiceId
            };

            return (GetInvoiceProductsFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> with existing <c>Product</c> with default price.
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>UoM</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid invoiceId, Guid productId, Guid uomId, decimal quantity)
        {
            return Add(invoiceId, productId, uomId, quantity, null, null, null);
        }

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> with existing <c>Product</c> with overridden price.
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>UoM</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">Override the <c>Product</c> 's price</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid invoiceId, Guid productId, Guid uomId, decimal quantity, decimal price)
        {
            return Add(invoiceId, productId, uomId, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> with existing <c>Product</c>.
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>UoM</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">If you want to override the <c>Product</c> 's price, otherwise the <c>Product</c> 's default price will be used for calculations</param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid invoiceId, Guid productId, Guid uomId, decimal quantity, decimal? price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(invoiceId, "invoiceId");
            ExceptionThrow.IfGuidEmpty(productId, "productId");
            ExceptionThrow.IfGuidEmpty(uomId, "uomId");
            ExceptionThrow.IfNegative(quantity, "quantity");

            var entity = XProduct.Create(this.EntityName, "invoice", invoiceId, productId, uomId, "", quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> with manual product information.
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid invoiceId, string productName, decimal quantity, decimal price)
        {
            return Add(invoiceId, productName, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Invoice product</c> to <c>Invoice</c> with manual product information.
        /// </summary>
        /// <param name="invoiceId"><c>Invoice</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid invoiceId, string productName, decimal quantity, decimal price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(invoiceId, "invoiceId");
            ExceptionThrow.IfNullOrEmpty(productName, "productName");
            ExceptionThrow.IfNegative(quantity, "quantity");
            ExceptionThrow.IfNegative(price, "price");

            var entity = XProduct.Create(this.EntityName, "invoice", invoiceId, null, null, productName, quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        #endregion
    }
}
