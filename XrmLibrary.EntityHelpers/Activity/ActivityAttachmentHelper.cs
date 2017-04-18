using System;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;
using XrmLibrary.EntityHelpers.Utils;

namespace XrmLibrary.EntityHelpers.Activity
{
    /// <summary>
    /// An <c>Activity Mime Attachment</c> represents an attachment to an email message or an email template.
    /// This class provides mostly used common methods for <c>ActivityMimeAttachment</c> entity
    /// <para> 
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309364(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class ActivityAttachmentHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public ActivityAttachmentHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309364(v=crm.8).aspx";
            this.EntityName = "activitymimeattachment";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create an attachment from drive path.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg328344(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="emailId"><c>Email Activity</c> Id</param>
        /// <param name="path">File path</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Attach(Guid emailId, string path)
        {
            ExceptionThrow.IfGuidEmpty(emailId, "emailId");
            ExceptionThrow.IfNullOrEmpty(path, "path");

            var attachment = AttachmentHelper.CreateFromPath(path, string.Empty);
            return Attach(emailId, attachment.Data.ByteArray, attachment.Meta.Name);
        }

        /// <summary>
        /// Create an attachment from byte array.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg328344(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="emailId"><c>Email Activity</c> Id</param>
        /// <param name="content">File content</param>
        /// <param name="fileName">This will be attachment name</param>
        /// <returns>
        /// Created record Id (<see cref="Guid"/>)
        /// </returns>
        public Guid Attach(Guid emailId, byte[] content, string fileName)
        {
            ExceptionThrow.IfGuidEmpty(emailId, "emailId");
            ExceptionThrow.IfNull(content, "content");
            ExceptionThrow.IfNegative(content.Length, "content");
            ExceptionThrow.IfNullOrEmpty(fileName, "fileName");

            var entity = PrepareAttachmentData(emailId, content, fileName);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Create an attachment from <see cref="Stream"/>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg328344(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="emailId">Existing <c>Email Activity</c> Id</param>
        /// <param name="stream">File content</param>
        /// <param name="fileName">This will be attachment name</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Attach(Guid emailId, Stream stream, string fileName)
        {
            ExceptionThrow.IfGuidEmpty(emailId, "emailId");
            ExceptionThrow.IfNull(stream, "stream");
            ExceptionThrow.IfNullOrEmpty(fileName, "fileName");

            return Attach(emailId, stream.ToByteArray(), fileName);
        }

        /// <summary>
        /// Create an email attachment from <see cref="AttachmentData"/> object
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg328344(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="emailId">Existing <c>Email Activity</c> Id</param>
        /// <param name="attachmentFile"><see cref="AttachmentData"/> object</param>
        /// <returns>Created record Id (<see cref="System.Guid"/>)</returns>
        public Guid Attach(Guid emailId, AttachmentData attachmentFile)
        {
            ExceptionThrow.IfGuidEmpty(emailId, "emailId");
            ExceptionThrow.IfNull(attachmentFile, "attachmentFile");

            return Attach(emailId, attachmentFile.Data.ByteArray, attachmentFile.Meta.Name);
        }

        /// <summary>
        /// Attach an existing attachment to email
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg328344(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="emailId">Existing <c>Email Activity</c> Id</param>
        /// <param name="attachmentId">Existing <c>Attachment</c> Id</param>
        /// <returns>Created record Id (<see cref="System.Guid"/>)</returns>
        public Guid Attach(Guid emailId, Guid attachmentId)
        {
            ExceptionThrow.IfGuidEmpty(emailId, "EmailId");
            ExceptionThrow.IfGuidEmpty(attachmentId, "attachmentId");

            Entity existing = this.OrganizationService.Retrieve(this.EntityName, attachmentId, new ColumnSet(true));
            ExceptionThrow.IfNull(existing, "Attachment");

            Entity entity = new Entity(this.EntityName);
            entity["objectid"] = new EntityReference("email", emailId);
            entity["objecttypecode"] = "email";
            entity["attachmentid"] = existing["attachmentid"];

            return this.OrganizationService.Create(entity);
        }

        #endregion

        #region | Private Methods |

        Entity PrepareAttachmentData(Guid emailId, byte[] content, string fileName)
        {
            Entity result = new Entity(this.EntityName);
            result["objectid"] = new EntityReference("email", emailId);
            result["objecttypecode"] = "email";
            result["body"] = content.ToBase64String();
            result["filename"] = fileName;

            return result;
        }

        #endregion
    }
}
