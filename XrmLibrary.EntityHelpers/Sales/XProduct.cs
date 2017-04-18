using System;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;

namespace XrmLibrary.EntityHelpers.Sales
{
    class XProduct
    {
        #region | Public Methods |

        public static Entity Create(string entityName, string relatedEntityName, Guid relatedEntityId, Guid? productId, Guid? uomId, string productName, decimal quantity, decimal? price, decimal? discountAmount, decimal? tax)
        {
            ExceptionThrow.IfNullOrEmpty(entityName, "entityName");
            ExceptionThrow.IfNullOrEmpty(relatedEntityName, "relatedEntityName");
            ExceptionThrow.IfGuidEmpty(relatedEntityId, "relatedEntityId");
            ExceptionThrow.IfNegative(quantity, "quantity");

            Entity entity = new Entity(entityName.ToLower().Trim());

            entity[string.Format("{0}id", relatedEntityName.Trim())] = new EntityReference(relatedEntityName.Trim(), relatedEntityId);
            entity["quantity"] = quantity;

            if (string.IsNullOrEmpty(productName))
            {
                ExceptionThrow.IfNull(productId, "productId");
                ExceptionThrow.IfNull(uomId, "uomId");
                ExceptionThrow.IfGuidEmpty(productId.Value, "productId");
                ExceptionThrow.IfGuidEmpty(uomId.Value, "uomId");

                entity["productid"] = new EntityReference("product", productId.Value);
                entity["uomid"] = new EntityReference("uom", uomId.Value);
            }
            else
            {
                entity["isproductoverridden"] = true;
                entity["productdescription"] = productName;
            }

            if (price.HasValue)
            {
                ExceptionThrow.IfNegative(price.Value, "price");

                entity["priceperunit"] = new Money(price.Value);
                entity["ispriceoverridden"] = true;
            }

            if (discountAmount.HasValue)
            {
                ExceptionThrow.IfNegative(discountAmount.Value, "discountAmount");

                entity["manualdiscountamount"] = new Money(discountAmount.Value);
            }

            if (tax.HasValue)
            {
                ExceptionThrow.IfNegative(tax.Value, "tax");

                entity["tax"] = new Money(tax.Value);
            }

            return entity;
        }

        #endregion
    }
}
