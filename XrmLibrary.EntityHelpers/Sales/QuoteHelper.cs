using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Sales
{
    /// <summary>
    /// A <c>Quote</c> is a formal offer for products and/or services, proposed at specific prices and related payment terms that is sent to a prospective customer. 
    /// This class provides mostly used common methods for <c>Quote</c> entity
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg328120(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class QuoteHelper : BaseEntityHelper
    {
        #region | Known Issues |

        /*
         * KNOWNISSUE : Quote.ReviseQuoteRequest : 
         * Quote record have to be "closed" before call this action. Otherwise SDK throws an exception "The quote cannot be revised because it is not in closed state."
         * 
         */

        #endregion

        #region | Private Definitions |

        bool _useUtc = false;
        string[] _quoteColumns = new string[] { "quoteid", "name", "quotenumber", "revisionnumber" };
        string[] _salesorderColumns = new string[] { "salesorderid", "totalamount", "ordernumber" };

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Quote</c> 's <c>Draft</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Quote
        /// </para>
        /// </summary>
        public enum QuoteDraftStatusCode
        {
            CustomStatusCode = 0,
            InProgress = 1
        }

        /// <summary>
        /// <c>Quote</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Quote
        /// </para>
        /// </summary>
        public enum QuoteActiveStatusCode
        {
            CustomStatusCode = 0,
            InProgress = 2,
            Open = 3
        }

        /// <summary>
        /// <c>Quote</c> 's <c>Won</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Quote
        /// </para>
        /// </summary>
        public enum QuoteWonStatusCode
        {
            CustomStatusCode = 0,
            Won = 4
        }

        /// <summary>
        /// <c>Quote</c> 's <c>Closed</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Quote
        /// </para>
        /// </summary>
        public enum QuoteClosedStatusCode
        {
            CustomStatusCode = 0,
            Lost = 5,
            Cancelled = 6,
            Revised = 7
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        /// <param name="useUTC">
        /// Set <c>true</c>, if you want to use Utc datetime format, otherwise set <c>false</c>.
        /// This parameter's default value is <c>true</c>
        /// </param>
        public QuoteHelper(IOrganizationService service, bool useUTC = true) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg328120(v=crm.8).aspx";
            this.EntityName = "quote";
            this._useUtc = useUTC;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Change <c>Quote</c> status to <c>Active</c>.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="status"><see cref="QuoteActiveStatusCode"/> status code</param>
        /// <param name="customStatusCode"></param>
        /// <returns></returns>
        public void Activate(Guid id, QuoteActiveStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == QuoteActiveStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "CustomStatusCode");
                statusCode = customStatusCode;
            }

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, 1, statusCode);
        }

        /// <summary>
        /// Change <c>Quote</c> status to <c>Closed</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.closequoterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Quote</c> Id</param>
        /// <param name="status"><see cref="QuoteClosedStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="quotecloseSubject"><c>QuoteClose</c> subject</param>
        /// <param name="quotecloseDescription"><c>QuoteClose</c> description</param>
        /// <param name="quotecloseActualEnd"><c>QuoteClose</c> actual end date</param>
        /// <returns><see cref="CloseQuoteResponse"/></returns>
        public CloseQuoteResponse Close(Guid id, QuoteClosedStatusCode status, int customStatusCode = 0, string quotecloseSubject = "", string quotecloseDescription = "", DateTime? quotecloseActualEnd = null)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == QuoteClosedStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "CustomStatusCode");
                statusCode = customStatusCode;
            }

            var quoteEntity = this.Get(id, "quotenumber");

            if (string.IsNullOrEmpty(quotecloseSubject))
            {
                switch (status)
                {
                    case QuoteClosedStatusCode.Lost:
                        quotecloseSubject = string.Format("Quote Close (Lost) - {0}", quoteEntity.GetAttributeValue<string>("quotenumber"));
                        break;
                    case QuoteClosedStatusCode.Cancelled:
                        quotecloseSubject = string.Format("Quote Close (Cancelled) - {0}", quoteEntity.GetAttributeValue<string>("quotenumber"));
                        break;
                    case QuoteClosedStatusCode.Revised:
                        quotecloseSubject = string.Format("Quote Close (Revised) - {0}", quoteEntity.GetAttributeValue<string>("quotenumber"));
                        break;
                }
            }

            CloseQuoteRequest request = new CloseQuoteRequest()
            {
                QuoteClose = PrepareQuoteCloseEntity(id, quotecloseSubject, quotecloseDescription, quotecloseActualEnd),
                Status = new OptionSetValue(statusCode)
            };

            return (CloseQuoteResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Change <c>Quote</c> status to <c>Won</c>.
        /// Please note that this action is required in order to convert a <c>Quote</c> into a <c>SalesOrder</c> with <see cref="ConvertToSalesOrder(Guid, string[])"/> method.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.winquoterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Quote</c> Id</param>
        /// <param name="status"><see cref="QuoteWonStatusCode"/> status code</param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <param name="quotecloseSubject"><c>QuoteClose</c> subject</param>
        /// <param name="quotecloseDescription"><c>QuoteClose</c> description</param>
        /// <param name="quotecloseActualEnd"><c>QuoteClose</c> actual end date</param>
        /// <returns><see cref="WinQuoteResponse"/></returns>
        public WinQuoteResponse Win(Guid id, QuoteWonStatusCode status, int customStatusCode = 0, string quotecloseSubject = "", string quotecloseDescription = "", DateTime? quotecloseActualEnd = null)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            int statusCode = (int)status;

            if (status == QuoteWonStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            if (string.IsNullOrEmpty(quotecloseSubject))
            {
                var quoteEntity = this.Get(id, "quotenumber");
                quotecloseSubject = string.Format("Quote Close (Won) - {0}", quoteEntity.GetAttributeValue<string>("quotenumber"));
            }

            WinQuoteRequest request = new WinQuoteRequest()
            {
                QuoteClose = PrepareQuoteCloseEntity(id, quotecloseSubject, quotecloseDescription, quotecloseActualEnd),
                Status = new OptionSetValue(statusCode)
            };

            return (WinQuoteResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Close existing <c>Quote</c> as <c>Revised</c> status and create a new <c>Quote</c> as <c>Draft</c> status.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.revisequoterequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id">Existing <c>Quote</c> Id</param>
        /// <param name="closeBeforeRevise">
        /// Set this <c>true</c> if your <c>Quote</c> record is currently <c>Active</c> status.
        /// Otherwise SDK throws an exception "The quote cannot be revised because it is not in closed state."
        /// </param>
        /// <param name="retrieveColumns">
        /// Default attributes are "quoteid", "name", "quotenumber", "revisionnumber".
        /// If you need more or different attributes please set this parameter
        /// </param>
        /// <returns>
        /// Returns revised /newly created <c>Queote</c> in <see cref="ReviseQuoteResponse.Entity"/> property with attributes that defined in <c>retrievedQuoteColums</c> parameter
        /// </returns>
        public ReviseQuoteResponse Revise(Guid id, bool closeBeforeRevise, params string[] retrieveColumns)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            string[] columns = !retrieveColumns.IsNullOrEmpty() ? retrieveColumns : _quoteColumns;

            if (closeBeforeRevise)
            {
                /*
                 * INFO : We have to "Close" as revised current quote before revised.
                 * Otherwise SDK throws an exception "The quote cannot be revised because it is not in closed state."
                 */

                Close(id, QuoteClosedStatusCode.Revised, 0, "", "", null);
            }

            ReviseQuoteRequest request = new ReviseQuoteRequest()
            {
                QuoteId = id,
                ColumnSet = new ColumnSet(columns)
            };

            return (ReviseQuoteResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Create a <c>Quote</c> from an <c>Opportunity</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.generatequotefromopportunityrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="opportunityId"><c>Opportunity</c> Id</param>
        /// <param name="retrievedColumns">
        /// Default attributes are "quoteid", "name", "quotenumber", "revisionnumber".
        /// If you need more or different attributes please set this parameter
        /// </param>
        /// <returns>
        /// Returns created <c>Quote</c> in <see cref="GenerateQuoteFromOpportunityResponse.Entity"/> property with attributes that defined in <c>retrievedQuoteColums</c> parameter
        /// </returns>
        public GenerateQuoteFromOpportunityResponse CreateFromOpportunity(Guid opportunityId, params string[] retrievedColumns)
        {
            ExceptionThrow.IfGuidEmpty(opportunityId, "opportunityId");

            string[] columns = !retrievedColumns.IsNullOrEmpty() ? retrievedColumns : _quoteColumns;

            GenerateQuoteFromOpportunityRequest request = new GenerateQuoteFromOpportunityRequest
            {
                OpportunityId = opportunityId,
                ColumnSet = new ColumnSet(columns)
            };

            return (GenerateQuoteFromOpportunityResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Convert <c>Quote</c> (that is <c>Won</c> status) to a <c>SalesOrder</c>.
        /// <para>
        /// Please note that, <c>Quote</c> must be <c>Won</c> status to convert <c>SalesOrder</c>
        /// </para>
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.convertquotetosalesorderrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Quote</c> Id</param>
        /// <param name="retrievedColumns">
        /// Default attributes are "salesorderid", "totalamount", "ordernumber".
        /// If you need more or different attributes please set this parameter.
        /// </param>
        /// <returns>
        /// Returns created <c>Sales Order</c> in <see cref="ConvertQuoteToSalesOrderResponse.Entity"/> property with attributes that defined in <c>retrievedQuoteColums</c> parameter.
        /// </returns>
        public ConvertQuoteToSalesOrderResponse ConvertToSalesOrder(Guid id, params string[] retrievedColumns)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            string[] columns = !retrievedColumns.IsNullOrEmpty() ? retrievedColumns : _salesorderColumns;

            ConvertQuoteToSalesOrderRequest request = new ConvertQuoteToSalesOrderRequest()
            {
                QuoteId = id,
                ColumnSet = new ColumnSet(columns)
            };

            return (ConvertQuoteToSalesOrderResponse)this.OrganizationService.Execute(request);
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Create a <c>quoteclose</c> entity.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="subject"></param>
        /// <param name="description"></param>
        /// <param name="actualEnd"></param>
        /// <returns></returns>
        Entity PrepareQuoteCloseEntity(Guid id, string subject, string description, DateTime? actualEnd = null)
        {
            DateTime now = this._useUtc ? DateTime.UtcNow : DateTime.Now;

            Entity quoteClose = new Entity("quoteclose");
            quoteClose["quoteid"] = new EntityReference("quote", id);
            quoteClose["subject"] = subject;
            quoteClose["description"] = description;
            quoteClose["actualend"] = actualEnd.HasValue ? actualEnd.Value : now;

            return quoteClose;
        }

        #endregion
    }
}
