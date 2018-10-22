using Microsoft.Xrm.Sdk;

namespace XrmLibrary.ExtensionMethods
{
    public static class DecimalExtensions
    {
        public static Money ToMoney(this decimal decimalValue)
        {
            return new Money(decimalValue);
        }
    }
}
