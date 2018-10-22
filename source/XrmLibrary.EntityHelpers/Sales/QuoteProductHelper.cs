using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// The <c>Quote Product (detail)</c> entity stores a line item in a <c>Quote</c>. Each quote can have multiple quote detail (quote product) records.
    /// This class provides mostly used common methods for <c>QuoteDetail</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328283(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class QuoteProductHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public QuoteProductHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328120(v=crm.8).aspx";
            this.EntityName = "quotedetail";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> from an <c>Opportunity</c>.
        /// This retrieves the products from an <c>Opportunity</c> and copies them to the <c>Quote</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.getquoteproductsfromopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <returns><see cref="GetQuoteProductsFromOpportunityResponse"/></returns>
        public GetQuoteProductsFromOpportunityResponse Add(Guid quoteId, Guid opportunityId)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");
            ExceptionThrow.IfGuidEmpty(quoteId, "quoteId");

            GetQuoteProductsFromOpportunityRequest request = new GetQuoteProductsFromOpportunityRequest()
            {
                OpportunityId = opportunityId,
                QuoteId = quoteId
            };

            return (GetQuoteProductsFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> with existing <c>Product</c> with default price.
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid quoteId, Guid productId, Guid uomId, decimal quantity)
        {
            return Add(quoteId, productId, uomId, quantity, null, null, null);
        }

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> with existing <c>Product</c> with overridden price.
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>UoM</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">Override the <c>Product</c> 's price</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid quoteId, Guid productId, Guid uomId, decimal quantity, decimal price)
        {
            return Add(quoteId, productId, uomId, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> with existing <c>Product</c>.
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>UoM</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid quoteId, Guid productId, Guid uomId, decimal quantity, decimal? price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(quoteId, "quoteId");
            ExceptionThrow.IfGuidEmpty(productId, "productId");
            ExceptionThrow.IfGuidEmpty(uomId, "uomId");
            ExceptionThrow.IfNegative(quantity, "quantity");

            var entity = XProduct.Create(this.EntityName, "quote", quoteId, productId, uomId, "", quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> with manual product information.
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid quoteId, string productName, decimal quantity, decimal price)
        {
            return Add(quoteId, productName, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Quote product</c> to <c>Quote</c> with manual product information.
        /// </summary>
        /// <param name="quoteId"><c>Quote</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid quoteId, string productName, decimal quantity, decimal price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(quoteId, "quoteId");
            ExceptionThrow.IfNullOrEmpty(productName, "productName");
            ExceptionThrow.IfNegative(quantity, "quantity");
            ExceptionThrow.IfNegative(price, "price");

            var entity = XProduct.Create(this.EntityName, "quote", quoteId, null, null, productName, quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        #endregion
    }
}
