using System;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// An <c>Opportunity Product</c> represents an association between an <c>Opportunity</c> and a <c>Product</c>.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328362(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class OpportunityProductHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"></param>
        public OpportunityProductHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328362(v=crm.8).aspx";
            this.EntityName = "opportunityproduct";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>Opportunity product</c> to <c>Opportunity</c> with existing <c>Product</c> with default price.
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid opportunityId, Guid productId, Guid uomId, decimal quantity)
        {
            return Add(opportunityId, productId, uomId, quantity, null, null, null);
        }

        /// <summary>
        /// Add a <c>Opportunity product</c> to <c>Opportunity</c> with existing <c>Product</c> with overridden price.
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">Override the <c>Product</c> 's price
        /// </param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid opportunityId, Guid productId, Guid uomId, decimal quantity, decimal price)
        {
            return Add(opportunityId, productId, uomId, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Opportunity product</c> to <c>Opportunity</c> with existing <c>Product</c>.
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">
        /// If you want to override the <c>Product</c> 's price, otherwise the <c>Product</c> 's default price will be used for calculations
        /// </param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid opportunityId, Guid productId, Guid uomId, decimal quantity, decimal? price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");
            ExceptionThrow.IfGuidEmpty(productId, "productId");
            ExceptionThrow.IfGuidEmpty(uomId, "uomId");
            ExceptionThrow.IfNegative(quantity, "quantity");

            var entity = XProduct.Create(this.EntityName, "opportunity", opportunityId, productId, uomId, "", quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add a <c>Opportunity product</c> to <c>Opportunity</c> with manual product information.
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid opportunityId, string productName, decimal quantity, decimal price)
        {
            return Add(opportunityId, productName, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Opportunity product</c> to <c>Opportunity</c> with manual product information.
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid opportunityId, string productName, decimal quantity, decimal price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");
            ExceptionThrow.IfNullOrEmpty(productName, "productName");
            ExceptionThrow.IfNegative(quantity, "quantity");
            ExceptionThrow.IfNegative(price, "price");

            var entity = XProduct.Create(this.EntityName, "opportunity", opportunityId, null, null, productName, quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        #endregion
    }
}
