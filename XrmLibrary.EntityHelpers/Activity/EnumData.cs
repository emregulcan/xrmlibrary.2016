using System.ComponentModel;

namespace XrmLibrary.EntityHelpers.Activity
{
    #region | Enums |

    /// <summary>
    /// An <c>Activity Party</c> represents a person or group associated with an activity. An activity can have multiple activity parties.
    /// There are 11 types of <c>ActivityParty</c> in Microsoft Dynamics CRM (2015).
    /// The activity party type is stored as an integer value in the <c>ActivityParty.ParticipationTypeMask</c> attribute.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328549(v=crm.8).aspx
    /// </para>
    /// </summary>
    public enum ActivityPartyType
    {
        /// <summary>
        /// Undefined - FOR INTERNAL USE ONLY.
        /// Please do not use this.
        /// </summary>
        Undefined = 0,

        /// <summary>
        /// Specifies the sender.
        /// </summary>
        Sender = 1,

        /// <summary>
        /// Specifies the recipient in the To field.
        /// </summary>
        ToRecipient = 2,

        /// <summary>
        /// Specifies the recipient in the Cc field.
        /// </summary>
        CCRecipient = 3,

        /// <summary>
        /// Specifies the recipient in the Bcc field.
        /// </summary>
        BccRecipient = 4,

        /// <summary>
        /// Specifies a required attendee.
        /// </summary>
        RequiredAttendee = 5,

        /// <summary>
        /// Specifies an optional attendee.
        /// </summary>
        OptionalAttendee = 6,

        /// <summary>
        /// Specifies the activity organizer.
        /// </summary>
        Organizer = 7,

        /// <summary>
        /// Specifies the regarding item.
        /// </summary>
        Regarding = 8,

        /// <summary>
        /// Specifies the activity owner.
        /// </summary>
        Owner = 9,

        /// <summary>
        /// Specifies a resource.
        /// </summary>
        Resource = 10,

        /// <summary>
        /// Specifies a customer.
        /// </summary>
        Customer = 11
    }

    /// <summary>
    /// <c>Email</c> activity 's <c>Active</c> statuscode values
    /// <para>
    /// For more information look at https://technet.microsoft.com/en-us/library/dn531157(v=crm.8).aspx#BKMK_Email
    /// </para>
    /// </summary>
    public enum EmailActiveStatusCode
    {
        CustomStatusCode = 0,
        Draft = 1,
        Failed = 8
    }

    /// <summary>
    /// <c>Email</c> activity 's <c>Completed</c> statuscode values
    /// <para>
    /// For more information look at https://technet.microsoft.com/en-us/library/dn531157(v=crm.8).aspx#BKMK_Email
    /// </para>
    /// </summary>
    public enum EmailCompletedStatusCode
    {
        CustomStatusCode = 0,
        Completed = 2,
        Sent = 3,
        Received = 4,
        PendingSend = 6,
        Sending = 7
    }

    /// <summary>
    /// <c>Email</c> activity 's <c>Cancelled</c> statuscode values
    /// <para>
    /// For more information look at https://technet.microsoft.com/en-us/library/dn531157(v=crm.8).aspx#BKMK_Email
    /// </para>
    /// </summary>
    public enum EmailCanceledStatusCode
    {
        CustomStatusCode = 0,
        Cancelled = 5
    }

    /// <summary>
    /// <c>Email</c> activity 's <c>Priority</c> values
    /// </summary>
    public enum PriorityCode
    {
        Low = 0,
        Normal = 1,
        High = 2
    }

    public enum DirectionCode
    {
        [Description("false")]
        Incoming,

        [Description("true")]
        Outgoing
    }

    #endregion
}
