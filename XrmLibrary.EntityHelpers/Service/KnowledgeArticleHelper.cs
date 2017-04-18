using System;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Service
{
    /// <summary>
    /// A <c>Knowledge Article</c> is a type of structured content that is managed internally as part of a knowledge base. 
    /// With the <c>Knowledge Article</c>, you can manage the distribution of product and service information for a business unit. 
    /// This class provides mostly used common methods for <c>Knowledge Article</c> entity.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309378(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class KnowledgeArticleHelper : BaseEntityHelper
    {
        #region | Bookmarks |

        /*
         * INFO : Bookmarks.KnowledgeArticle
         * https://msdn.microsoft.com/en-us/library/gg309345.aspx
         * http://crmbook.powerobjects.com/basics/service-management-overview/knowledge-base/
         * https://community.dynamics.com/crm/b/crmcat/archive/2015/07/12/how-to-use-articles-in-microsoft-dynamics-crm-amp-why-you-should-use-them 
         * https://blogs.msdn.microsoft.com/crm/2015/10/15/new-knowledge-management-in-microsoft-dynamics-crm-2016-release/
         * http://www.inogic.com/blog/2016/07/insights-for-basics-of-knowledge-articles-in-dynamics-crm-2016
         * https://www.microsoft.com/en-us/dynamics/crm-customer-center/find-knowledge-articles-from-within-a-record-in-dynamics-365.aspx
         * http://blog.sonomapartners.com/2016/01/crm-2016-knowledge-articles-and-full-text-knowledge-search.html
         * 
         */

        #endregion

        #region | Known Issues |

        /*
         * KNOWNISSUE : KnowledgeArticleHelper.StatusCode / StateCode
         * With CRM Online 2016 / D365 KnowledArticle entity released.
         * On 2017-04-04 still old State - StatusReason codes at https://technet.microsoft.com/en-us/library/dn531157.aspx
         * For more information look at "Knowledge article lifecycle: Change the state of a knowledge article" https://msdn.microsoft.com/en-us/library/gg309345.aspx 
         * Following State - StatusReason codes in this class, copied from D365 Online customization.
         * 
         * 
         * KNOWNISSUE : KnowledgebaseHelper.Search
         * You must define "PagingInfo" for "QueryExpression".
         * For more information look at http://blog.sonomapartners.com/2016/01/crm-2016-knowledge-articles-and-full-text-knowledge-search.html
         */

        #endregion

        #region | Notes |

        /*
         * INFO : Knowledgebase CRM2016 - D365 Notes
         * 
         * https://msdn.microsoft.com/en-us/library/gg309378(v=crm.8).aspx
         * With CRM Online 2016 Update 1 and Microsoft Dynamics CRM 2016 Service Pack 1 (on-premises) release, the following entities used for knowledge management are deprecated: KbArticle, KbArticleComment, and KbArticleTemplate. 
         * These entities won't be supported in a future major release of Dynamics 365. 
         * You must use the newer "KnowledgeArticle" entity in your code for knowledge management in Dynamics 365. 
         */

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Knowledge Article</c> search type
        /// </summary>
        public enum SearchTypeCode
        {
            FullText = 4
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's statecode values
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.fulltextsearchknowledgearticlerequest.statecode.aspx
        /// </para>
        /// </summary>
        public enum KnowledgeArticleStateCode
        {
            [Description("draft")]
            Draft = 0,

            [Description("approved")]
            Approved = 1,

            [Description("scheduled")]
            Scheduled = 2,

            [Description("published")]
            Published = 3,

            [Description("expired")]
            Expired = 4,

            [Description("archived")]
            Archived = 5,

            [Description("discarded")]
            Discarded = 6
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Draft</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleDraftStatusCode
        {
            CustomStatusCode = 0,
            Proposed = 1,
            Draft = 2,
            NeedsReview = 3,
            InReview = 4
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Approved</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleApprovedStatusCode
        {
            CustomStatusCode = 0,
            Approved = 5
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Scheduled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleScheduledStatusCode
        {
            CustomStatusCode = 0,
            Scheduled = 6
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Published</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticlePublishedStatusCode
        {
            CustomStatusCode = 0,
            Published = 7,
            NeedsReview = 8,
            Updating = 9
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Expired</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleExpiredStatusCode
        {
            CustomStatusCode = 0,
            Expired = 10,
            Rejected = 11
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Archived</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleArchivedStatusCode
        {
            CustomStatusCode = 0,
            Archived = 12
        }

        /// <summary>
        /// <c>Knowledge Article</c> 's <c>Discarded</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Article
        /// </para>
        /// </summary>
        public enum KnowledgeArticleDiscardedStatusCode
        {
            CustomStatusCode = 0,
            Discarded = 13,
            Rejected = 14
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public KnowledgeArticleHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309378(v=crm.8).aspx";
            this.EntityName = "knowledgearticle";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Approve <c>Knowledge Article</c> record.
        /// </summary>
        /// <param name="id"><c>Article</c> Id</param>
        public void Approve(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)KnowledgeArticleStateCode.Approved, (int)KnowledgeArticleApprovedStatusCode.Approved);
        }

        /// <summary>
        /// Publish <c>Knowledge Article</c> record.
        /// <para>
        /// Please note that if you <c>Publish</c> new version of existing <c>Knowledge Article</c>, main <c>Knowledge Article</c>'s status will be <c>Archived</c>.
        /// </para>
        /// </summary>
        /// <param name="id"><c>Article</c> Id</param>
        public void Publish(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)KnowledgeArticleStateCode.Published, (int)KnowledgeArticlePublishedStatusCode.Published);
        }

        /// <summary>
        /// Revert to draft <c>Knowledge Article</c> record.
        /// </summary>
        /// <param name="id"><c>Article</c> Id</param>
        public void Draft(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)KnowledgeArticleStateCode.Draft, (int)KnowledgeArticleDraftStatusCode.Proposed);
        }

        /// <summary>
        /// Archive <c>Knowledge Article</c> record.
        /// </summary>
        /// <param name="id"><c>Article</c> Id</param>
        public void Archive(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)KnowledgeArticleStateCode.Archived, (int)KnowledgeArticleArchivedStatusCode.Archived);
        }

        /// <summary>
        /// Search for <c>Knowledge Article</c>.
        /// <para>
        /// For more information look at 
        /// <c>By Fulltext</c> : https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.fulltextsearchknowledgearticlerequest.aspx
        /// </para>
        /// </summary>
        /// <param name="searchType"><see cref="SearchTypeCode"/></param>
        /// <param name="query">
        /// Query criteria to find <c>Knowledge Article</c> with specified keyword.
        /// This parameter supports <c>QueryExpression</c> and <c>FetchXml</c>.
        /// <para>
        /// Please note that <see cref="PagingInfo"/> data must be defined in this query.
        /// </para>
        /// </param>
        /// <param name="searchText">Text to search for in <c>Knowledge Article</c> data</param>
        /// <param name="useInflection">Indicates whether to use inflectional stem matching when searching for knowledge base articles with a specified body text</param>
        /// <param name="removeDuplicates">
        /// Indicates whether to remove multiple versions of the same knowledge article in search results. 
        /// Default value is <c>true</c>.
        /// </param>
        /// <param name="stateCode">
        /// State of the knowledge articles to filter the search results.
        /// Default value is <c>-1</c> (for all data). For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.fulltextsearchknowledgearticlerequest.statecode.aspx
        /// </param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Article</c> data.
        /// </returns>
        public EntityCollection Search(SearchTypeCode searchType, QueryBase query, string searchText, bool useInflection, bool removeDuplicates = true, int stateCode = -1)
        {
            ExceptionThrow.IfNull(query, "query");
            ExceptionThrow.IfNullOrEmpty(searchText, "searchText");

            if (query is QueryExpression)
            {
                ExceptionThrow.IfNullOrEmpty(((QueryExpression)query).EntityName, "QueryExpression.EntityName");
                ExceptionThrow.IfNotExpectedValue(((QueryExpression)query).EntityName.ToLower(), "QueryExpression.EntityName", this.EntityName.ToLower());
            }

            EntityCollection result = new EntityCollection();

            OrganizationRequest serviceRequest = null;
            OrganizationResponse serviceResponse = null;

            switch (searchType)
            {
                case SearchTypeCode.FullText:
                    serviceRequest = new FullTextSearchKnowledgeArticleRequest()
                    {
                        QueryExpression = query,
                        RemoveDuplicates = removeDuplicates,
                        SearchText = searchText,
                        StateCode = stateCode,
                        UseInflection = useInflection
                    };

                    serviceResponse = this.OrganizationService.Execute(serviceRequest);
                    result = ((FullTextSearchKnowledgeArticleResponse)serviceResponse).EntityCollection;
                    break;
            }

            return result;
        }

        /// <summary>
        /// Create <c>Translation</c> of a <c>Knowledge Article</c>.
        /// Please note that this creates a new <c>Knowledge Article</c> record with the <c>Title</c>, <c>Content</c>, <c>Description</c> and <c>Keywords</c> copied from the source record to the new record
        /// and the language of the new record set to the one you specified in the <paramref name="languageId"/> parameter.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.createknowledgearticletranslationrequest.aspx
        /// </para>
        /// </summary>
        /// <param name="sourceId"><c>Knowledge Article</c> Id</param>
        /// <param name="isMajor">Indicate create a major or minor version of the knowledge article</param>
        /// <param name="languageId">
        /// <c>LanguageLocale</c> Id.
        /// You can also use <see cref="LanguageCodeList"/> enumaration to find related Language Id (<see cref="Guid"/>).
        /// <para>
        /// Please note that the GUID value of the <c>primary key (LanguageLocaleId)</c> for each language record in the <c>LanguageLocale</c> entity is the same across all Dynamics 365 organizations.
        /// For more information look at https://msdn.microsoft.com/en-us/library/gg309345.aspx#Translation
        /// </para>
        /// For more information about <c>LanguageLocale</c> entity look at https://msdn.microsoft.com/en-us/library/mt607524.aspx
        /// </param>
        /// <returns>
        /// Returns newly created <c>Knowledge Article</c> by <see cref="EntityReference"/> in <see cref="CreateKnowledgeArticleTranslationResponse.CreateKnowledgeArticleTranslation"/> property.
        /// </returns>
        public CreateKnowledgeArticleTranslationResponse Translate(Guid sourceId, bool isMajor, Guid languageId)
        {
            ExceptionThrow.IfGuidEmpty(sourceId, "sourceId");
            ExceptionThrow.IfGuidEmpty(languageId, "languageId");

            CreateKnowledgeArticleTranslationRequest request = new CreateKnowledgeArticleTranslationRequest()
            {
                IsMajor = isMajor,
                Language = new EntityReference("languagelocale ", languageId),
                Source = new EntityReference(this.EntityName, sourceId)
            };

            return (CreateKnowledgeArticleTranslationResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Create a major or minor version of a <c>Knowledge Article</c> record.
        /// <para>
        /// Please note that if you want use new version of <c>Knowledge Article</c>, you should call <see cref="Approve(Guid)"/> and <see cref="Publish(Guid)"/> methods.
        /// </para>
        /// </summary>
        /// <param name="sourceId"><c>Knowledge Article</c> Id</param>
        /// <param name="isMajor">Indicate create a major or minor version of the knowledge article</param>
        /// <returns>
        /// Returns version record created for a <c>Knowledge Article</c> by <see cref="EntityReference"/> in <see cref="CreateKnowledgeArticleVersionResponse.CreateKnowledgeArticleVersion"/> property.
        /// </returns>
        public CreateKnowledgeArticleVersionResponse CreateVersion(Guid sourceId, bool isMajor = false)
        {
            ExceptionThrow.IfGuidEmpty(sourceId, "sourceId");

            CreateKnowledgeArticleVersionRequest request = new CreateKnowledgeArticleVersionRequest()
            {
                Source = new EntityReference(this.EntityName, sourceId),
                IsMajor = isMajor
            };

            return (CreateKnowledgeArticleVersionResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
