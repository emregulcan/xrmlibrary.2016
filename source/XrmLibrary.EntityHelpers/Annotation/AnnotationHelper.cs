using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;
using XrmLibrary.EntityHelpers.Utils;

namespace XrmLibrary.EntityHelpers.Annotation
{
    /// <summary>
    /// An <c>Annotation (note)</c> is a text entry that can be associated with any entity in CRM. 
    /// However, you can associate annotations with only those custom entities that are created with the <c>HasNotes</c> property set to true. 
    /// This class provides mostly used common methods for <c>Annotation</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334398(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class AnnotationHelper : BaseEntityHelper
    {
        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public AnnotationHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334398(v=crm.8).aspx";
            this.EntityName = "annotation";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Create an <c>Annotation</c>.
        /// </summary>
        /// <param name="relatedObject">Related object that attach note</param>
        /// <param name="subject">Subject</param>
        /// <param name="text">Note text</param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Create(EntityReference relatedObject, string subject, string text)
        {
            ExceptionThrow.IfNull(relatedObject, "relatedObject");
            ExceptionThrow.IfGuidEmpty(relatedObject.Id, "relatedObject.Id");
            ExceptionThrow.IfNullOrEmpty(relatedObject.LogicalName, "relatedObject.LogicalName");
            ExceptionThrow.IfNullOrEmpty(text, "text");

            Entity entity = PrepareBasicAnnotation(relatedObject, subject, text);
            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Create an <c>Annotation</c> with <see cref="AttachmentData"/>.
        /// </summary>
        /// <param name="relatedObject">Related object that attach note</param>
        /// <param name="subject">Subject</param>
        /// <param name="text">Note text</param>
        /// <param name="attachment"><see cref="AttachmentData"/></param>
        /// <returns>Created record Id (<see cref="Guid"/>)</returns>
        public Guid Create(EntityReference relatedObject, string subject, string text, AttachmentData attachment)
        {
            ExceptionThrow.IfNull(relatedObject, "relatedObject");
            ExceptionThrow.IfGuidEmpty(relatedObject.Id, "relatedObject.Id");
            ExceptionThrow.IfNullOrEmpty(relatedObject.LogicalName, "relatedObject.LogicalName");
            ExceptionThrow.IfNull(attachment, "attachment");
            ExceptionThrow.IfNull(attachment.Data, "attachment.Data");
            ExceptionThrow.IfNull(attachment.Meta, "attachment.Meta");
            ExceptionThrow.IfNullOrEmpty(attachment.Data.Base64, "attachment.Data.Base64");
            ExceptionThrow.IfNullOrEmpty(attachment.Meta.Name, "attachment.Meta.Name");

            Entity entity = PrepareBasicAnnotation(relatedObject, subject, text);

            entity["documentbody"] = attachment.Data.Base64;
            entity["filename"] = attachment.Meta.Name;
            entity["filesize"] = attachment.Meta.Size;
            entity["mimetype"] = attachment.Meta.MimeType;

            return this.OrganizationService.Create(entity);
        }

        /// <summary>
        /// Add attachment to existing <c>Annotation</c>.
        /// </summary>
        /// <param name="id"><c>Annotation</c> Id</param>
        /// <param name="attachment">
        /// <see cref="AttachmentData"/>
        /// </param>
        public void AddAttachment(Guid id, AttachmentData attachment)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");
            ExceptionThrow.IfNull(attachment, "attachment");
            ExceptionThrow.IfNull(attachment.Data, "attachment.Data");
            ExceptionThrow.IfNull(attachment.Meta, "attachment.Meta");
            ExceptionThrow.IfNullOrEmpty(attachment.Data.Base64, "attachment.Data.Base64");
            ExceptionThrow.IfNullOrEmpty(attachment.Meta.Name, "attachment.Meta.Name");

            Entity entity = new Entity("annotation");
            entity["annotationid"] = id;
            entity["documentbody"] = attachment.Data.Base64;
            entity["filename"] = attachment.Meta.Name;
            entity["filesize"] = attachment.Meta.Size;
            entity["mimetype"] = attachment.Meta.MimeType;

            this.OrganizationService.Update(entity);
        }

        /// <summary>
        /// Retrieve all <c>Annotation</c> data by parent record.
        /// </summary>
        /// <param name="parentId">Parent record Id</param>
        /// <param name="retrievedColumns">
        /// Default is <c>ColumnSet(true)</c>.
        /// </param>
        /// <returns>
        /// <seealso cref="EntityCollection"/> for <c>Annotation</c> data
        /// </returns>
        public EntityCollection GetAllByParent(Guid parentId, params string[] retrievedColumns)
        {
            ExceptionThrow.IfGuidEmpty(parentId, "parentId");

            ColumnSet columset = retrievedColumns.Length > 0 ? new ColumnSet(retrievedColumns) : new ColumnSet(true);

            QueryByAttribute query = new QueryByAttribute("annotation");
            query.AddAttributeValue("objectid", parentId);
            query.ColumnSet = columset;

            return this.OrganizationService.RetrieveMultiple(query);
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Create <c>Annotation</c> with basic data.
        /// </summary>
        /// <param name="relatedObject"></param>
        /// <param name="subject"></param>
        /// <param name="text"></param>
        /// <returns>
        /// <see cref="Entity"/> for prepared <c>Annotation</c>
        /// </returns>
        Entity PrepareBasicAnnotation(EntityReference relatedObject, string subject, string text)
        {
            Entity entity = new Entity(this.EntityName);
            entity["subject"] = subject;
            entity["notetext"] = text;
            entity["objectid"] = relatedObject;

            return entity;
        }

        #endregion
    }
}
