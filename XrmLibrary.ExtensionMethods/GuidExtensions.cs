using System;
using Microsoft.Xrm.Sdk;

namespace XrmLibrary.ExtensionMethods
{
    public static class GuidExtensions
    {
        #region | Public Methods |

        public static EntityReference ToEntityReference(this Guid id, string entityName)
        {
            return new EntityReference(entityName, id);
        }

        #endregion
    }
}
