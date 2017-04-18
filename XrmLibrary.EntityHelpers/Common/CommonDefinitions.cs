using System.ComponentModel;

namespace XrmLibrary.EntityHelpers.Common
{
    #region | Enums |

    /// <summary>
    /// <c>Principal</c> type code values
    /// </summary>
    public enum PrincipalType
    {
        [Description("systemuser")]
        SystemUser = 8,

        [Description("team")]
        Team = 9
    }

    /// <summary>
    /// <c>To</c> activityparty values
    /// </summary>
    public enum ToEntityType
    {
        [Description("lead")]
        Lead,

        [Description("account")]
        Account,

        [Description("contact")]
        Contact,

        [Description("queue")]
        Queue,

        [Description("systemuser")]
        SystemUser
    }


    public enum FromEntityType
    {
        [Description("queue")]
        Queue,

        [Description("systemuser")]
        SystemUser
    }

    public enum CustomerType
    {
        [Description("account")]
        Account = 1,

        [Description("contact")]
        Contact = 2
    }

    #endregion
}
