using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// A <c>SalesOrder Detail (order product)</c> entity stores a line item in an order. Each order can have multiple sales order detail instances.
    /// This class provides mostly used common methods for <c>SalesOrderDetail</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg327957(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class SalesOrderProductHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public SalesOrderProductHelper(IOrganizationService service) : base(service)
        {
            this.EntityName = "salesorderdetail";
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg327957(v=crm.8).aspx";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> from an <c>Opportunity</c>.
        /// This retrieves the set of <c>Product</c> data for a <c>SalesOrder</c> from the specified <c>Opportunity</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.getsalesorderproductsfromopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <returns>
        /// <see cref="GetSalesOrderProductsFromOpportunityResponse"/>
        /// </returns>
        public GetSalesOrderProductsFromOpportunityResponse Add(Guid salesorderId, Guid opportunityId)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");
            ExceptionThrow.IfGuidEmpty(salesorderId, "salesorderId");

            GetSalesOrderProductsFromOpportunityRequest request = new GetSalesOrderProductsFromOpportunityRequest()
            {
                OpportunityId = opportunityId,
                SalesOrderId = salesorderId
            };

            return (GetSalesOrderProductsFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> with existing <c>Product</c> with default price.
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid salesorderId, Guid productId, Guid uomId, decimal quantity)
        {
            return Add(salesorderId, productId, uomId, quantity, null, null, null);
        }

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> with existing <c>Product</c> with overridden price.
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">Override the <c>Product</c> 's price</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid salesorderId, Guid productId, Guid uomId, decimal quantity, decimal price)
        {
            return Add(salesorderId, productId, uomId, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> with existing <c>Product</c>.
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="productId"><c>Product</c> Id</param>
        /// <param name="uomId"><c>Unit</c> Id</param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price">
        /// If you want to override the <c>Product</c> 's price, otherwise the <c>Product</c> 's default price will be used for calculations
        /// </param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid salesorderId, Guid productId, Guid uomId, decimal quantity, decimal? price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(salesorderId, "salesorderId");
            ExceptionThrow.IfGuidEmpty(productId, "productId");
            ExceptionThrow.IfGuidEmpty(uomId, "uomId");
            ExceptionThrow.IfNegative(quantity, "quantity");

            var entity = XProduct.Create(this.EntityName, "salesorder", salesorderId, productId, uomId, "", quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> with manual product information.
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid salesorderId, string productName, decimal quantity, decimal price)
        {
            return Add(salesorderId, productName, quantity, price, null, null);
        }

        /// <summary>
        /// Add a <c>Sales Order product</c> to <c>SalesOrder</c> with manual product information.
        /// </summary>
        /// <param name="salesorderId"><c>SalesOrder</c> Id</param>
        /// <param name="productName"><c>Product Name</c></param>
        /// <param name="quantity"><c>Quantity</c></param>
        /// <param name="price"><c>Price</c></param>
        /// <param name="discountAmount"><c>Discount amount</c> by currency</param>
        /// <param name="tax"><c>Tax amount</c>by currency</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Add(Guid salesorderId, string productName, decimal quantity, decimal price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfGuidEmpty(salesorderId, "salesorderId");
            ExceptionThrow.IfNullOrEmpty(productName, "productName");
            ExceptionThrow.IfNegative(quantity, "quantity");
            ExceptionThrow.IfNegative(price, "price");

            var entity = XProduct.Create(this.EntityName, "salesorder", salesorderId, null, null, productName, quantity, price, discountAmount, tax);
            return this.OrganizationService.Create(entity);
        }

        #endregion
    }
}
