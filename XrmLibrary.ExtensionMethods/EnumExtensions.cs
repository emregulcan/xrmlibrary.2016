using System;
using Microsoft.Xrm.Sdk;

namespace XrmLibrary.ExtensionMethods
{
    public static class EnumExtensions
    {
        #region | Public Methods |

        public static OptionSetValue ToOptionSetValue(this Enum enumType)
        {
            return new OptionSetValue(Convert.ToInt32(enumType));
        }

        #endregion
    }
}
