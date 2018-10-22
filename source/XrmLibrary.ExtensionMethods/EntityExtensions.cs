using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace XrmLibrary.ExtensionMethods
{
    public static class EntityExtensions
    {
        #region | Public Methods |

        public static Dictionary<string, object> ToDictionary(this Entity entity)
        {
            return entity.Attributes.ToDictionary(k => k.Key, v => v.Value);
        }

        public static T GetValue<T>(this Entity entity, string attributeName)
        {
            T result = default(T);

            object entityAttribute = entity[attributeName];

            object resultValue = null;

            if (entityAttribute is EntityReference && typeof(T) == typeof(Guid))
            {
                resultValue = (entityAttribute as EntityReference).Id;
            }
            else if (entityAttribute is OptionSetValue && typeof(T) == typeof(int))
            {
                resultValue = (entityAttribute as OptionSetValue).Value;
            }
            if (entityAttribute == typeof(Money) && typeof(T) == typeof(decimal))
            {
                resultValue = (entityAttribute as Money).Value;
            }
            else
            {
                resultValue = entityAttribute;
            }

            result = (T)Convert.ChangeType(resultValue, typeof(T));

            return result;
        }

        public static T GetAliasedValue<T>(this Entity entity, string attributeName)
        {
            return entity.Contains(attributeName) ? (T)((AliasedValue)entity[attributeName]).Value : default(T);
        }

        public static string GetFormattedValue(this Entity entity, string attributeName)
        {
            return entity.FormattedValues.Contains(attributeName) ? entity.FormattedValues[attributeName] : string.Empty;
        }

        public static void SetValue(this Entity entity, string attribute, object value)
        {
            entity[attribute] = value;
        }

        public static Dictionary<string, EntityReference> GetAllEntityReference(this Entity entity)
        {
            return entity.Attributes.Where(e => e.Value is EntityReference).ToDictionary(k => k.Key, v => (EntityReference)v.Value);
        }

        public static Dictionary<string, OptionSetValue> GetAllOptionSetValue(this Entity entity)
        {
            return entity.Attributes.Where(e => e.Value is OptionSetValue).ToDictionary(k => k.Key, v => (OptionSetValue)v.Value);
        }

        public static Dictionary<string, Money> GetAllMoney(this Entity entity)
        {
            return entity.Attributes.Where(e => e.Value is Money).ToDictionary(k => k.Key, v => (Money)v.Value);
        }

        public static Entity Clone(this Entity entity, bool includePrimaryKey = false)
        {
            var result = new Entity(entity.LogicalName);
            KeyValuePair<string, object>[] attributeList = null;
            entity.Attributes.CopyTo(attributeList, 0);

            result.Attributes.AddRange(attributeList);

            if (result.Attributes.ContainsKey(entity.LogicalName + "id") && !includePrimaryKey)
            {
                result.Attributes.Remove(entity.LogicalName + "id");
            }

            return result;
        }

        public static Entity Clone(this Entity entity, string customPrefix, bool includePrimaryKey = false)
        {
            var result = new Entity(entity.LogicalName);
            string primaryKey = entity.LogicalName + "id";

            KeyValuePair<string, object>[] attributeList = null;
            entity.Attributes.CopyTo(attributeList, 0);

            result.Attributes.AddRange(attributeList.Where(a => a.Key.StartsWith(customPrefix)));

            if (result.Attributes.ContainsKey(primaryKey) && !includePrimaryKey)
            {
                result.Attributes.Remove(primaryKey);
            }
            else if (includePrimaryKey && entity.Contains(primaryKey))
            {
                result[primaryKey] = entity[primaryKey];
            }

            return result;
        }

        public static Entity Clone(this Entity entity, params string[] excludeColumns)
        {
            var result = new Entity(entity.LogicalName);
            KeyValuePair<string, object>[] attributeList = null;
            entity.Attributes.CopyTo(attributeList, 0);

            result.Attributes.AddRange(attributeList.Where(a => excludeColumns != null ? !excludeColumns.Contains(a.Key) : true));

            return result;
        }

        public static DataTable CreateDataTableSchema(this Entity entity)
        {
            var result = new DataTable(entity.LogicalName);
            entity.Attributes.Select(e => new DataColumn(e.Key, NativeType(e.Value))).ToList().ForEach(s => result.Columns.Add(s));

            return result;
        }

        private static Type NativeType(object attributeValue)
        {
            if (attributeValue is EntityReference)
            {
                return typeof(Guid);
            }
            else if (attributeValue is Money)
            {
                return typeof(decimal);
            }
            else if (attributeValue is OptionSetValue)
            {
                return typeof(int);
            }

            return attributeValue.GetType();
        }

        private static object NativeValue(object attributeValue)
        {
            if (attributeValue is EntityReference)
            {
                return (attributeValue as EntityReference).Id;
            }
            else if (attributeValue is Money)
            {
                return (attributeValue as Money).Value;
            }
            else if (attributeValue is OptionSetValue)
            {
                return (attributeValue as OptionSetValue).Value;
            }

            return attributeValue;
        }

        public static DataTable ToDataTable(this EntityCollection entityList)
        {
            var result = new DataTable(entityList.EntityName);

            if (entityList == null || entityList.Entities.Count == 0)
            {
                return result;
            }

            entityList.Entities[0].Attributes.Select(e => new DataColumn(e.Key, NativeType(e.Value))).ToList().ForEach(s => result.Columns.Add(s));
            foreach (var entity in entityList.Entities)
            {
                var row = result.NewRow();

                foreach (var attribute in entity.Attributes)
                {
                    row[attribute.Key] = NativeValue(attribute.Value);
                }

                result.Rows.Add(row);
            }

            return result;
        }

        public static bool IsNullOrEmpty(this EntityCollection entityList)
        {
            return entityList == null || entityList.Entities.Count == 0;
        } 

        #endregion
    }
}
