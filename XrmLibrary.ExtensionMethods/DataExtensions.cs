using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Xrm.Sdk;

namespace XrmLibrary.ExtensionMethods
{
    public static class DataExtensions
    {
        #region | Public Methods |

        public static Entity ToEntity(this DataRow dataRow)
        {
            Entity result = new Entity(dataRow.Table.TableName);

            foreach (DataColumn column in dataRow.Table.Columns)
            {
                result[column.ColumnName] = dataRow[column.ColumnName];
            }

            return result;
        }

        public static Entity ToEntity(this DataRow dataRow, string entityName)
        {
            dataRow.Table.TableName = entityName;
            return dataRow.ToEntity();
        }

        public static EntityCollection ToEntityCollection(this DataTable table)
        {
            EntityCollection result = new EntityCollection();
            result.EntityName = table.TableName;

            foreach (DataRow row in table.Rows)
            {
                result.Entities.Add(row.ToEntity());
            }

            return result;
        }

        public static EntityCollection ToEntityCollection(this DataTable table, string entityName)
        {
            table.TableName = entityName;
            return table.ToEntityCollection();
        }

        public static EntityCollection ToEntityCollection(this DataTable table, Dictionary<string, Type> columnTypeMap)
        {
            EntityCollection result = new EntityCollection();
            result.EntityName = table.TableName;

            foreach (DataRow row in table.Rows)
            {
                result.Entities.Add(row.ToEntity());
            }

            return result;
        }

        public static IList<Entity> ToEntityList(this DataTable table)
        {
            return table.ToEntityCollection().Entities;
        }

        #endregion

        #region | Private Methods |

        private static object GetAttributValue(object attributeValue, Type type)
        {
            if (type == typeof(EntityReference))
            {
                return new EntityReference(string.Empty, (Guid)attributeValue);
            }
            else if (type == typeof(Money))
            {
                return new Money((decimal)attributeValue);
            }
            else if (type == typeof(OptionSetValue))
            {
                return new OptionSetValue((int)attributeValue);
            }

            return attributeValue;
        }

        #endregion
    }
}
