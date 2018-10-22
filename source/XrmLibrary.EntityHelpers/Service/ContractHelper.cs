using System;
using System.ComponentModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.Service
{
    /// <summary>
    /// A <c>Contract</c> is an agreement to provide customer service support during specified coverage dates or for a specified number of cases or length of time. 
    /// The <c>Contract</c> entity is used to track customer service agreements. New contracts are created based on the <c>Contract Template</c>. You can create contracts only for existing accounts and contacts.
    /// This class provides mostly used common methods for <c>Contract</c> entity.
    /// <para>
    /// Please note that a <c>Contract</c> 's status is <c>Draft</c> until it is <c>Invoiced</c>.
    /// </para>
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg334373(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class ContractHelper : BaseEntityHelper
    {
        #region | Private Definitions |

        bool _useUtc = false;

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Contract</c> 's statecode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractStateCode
        {
            [Description("draft")]
            Draft = 0,

            [Description("invoiced")]
            Invoiced = 1,

            [Description("active")]
            Active = 2,

            [Description("onhold")]
            OnHold = 3,

            [Description("calceled")]
            Canceled = 4,

            [Description("expired")]
            Expired = 5
        }

        /// <summary>
        /// <c>Contract</c> 's <c>Draft</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractDraftStatusCode
        {
            CustomStatusCode = 0,
            Draft = 1
        }

        /// <summary>
        /// <c>Contract</c> 's <c>Invoiced</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractInvoicedStatusCode
        {
            CustomStatusCode = 0,
            Invoiced = 2
        }

        /// <summary>
        /// <c>Contract</c> 's <c>Active</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractActiveStatusCode
        {
            CustomStatusCode = 0,
            Active = 3
        }

        /// <summary>
        /// <c>Contract</c> 's <c>OnHold</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractOnHoldStatusCode
        {
            CustomStatusCode = 0,
            OnHold = 4
        }

        /// <summary>
        /// <c>Contract</c> 's <c>Canceled</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractCanceledStatusCode
        {
            CustomStatusCode = 0,
            Canceled = 5
        }

        /// <summary>
        /// <c>Contract</c> 's <c>Expired</c> statuscode values
        /// <para>
        /// For more information look at https://technet.microsoft.com/en-us/library/dn531157.aspx#BKMK_Contract
        /// </para>
        /// </summary>
        public enum ContractExpiredStatusCode
        {
            CustomStatusCode = 0,
            Expired = 6
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
        public ContractHelper(IOrganizationService service, bool useUTC = true) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg334373(v=crm.8).aspx";
            this.EntityName = "contract";
            this._useUtc = useUTC;
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Set status to <c>Invoiced</c>.
        /// </summary>
        /// <param name="id"><c>Conract</c> Id</param>
        public void Invoiced(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)ContractStateCode.Invoiced, (int)ContractInvoicedStatusCode.Invoiced);
        }

        /// <summary>
        /// Put the <c>Contract</c> to <c>Hold</c>.
        /// </summary>
        /// <param name="id"><c>Conract</c> Id</param>
        public void Hold(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)ContractStateCode.OnHold, (int)ContractOnHoldStatusCode.OnHold);
        }

        /// <summary>
        /// Release Hold.
        /// </summary>
        /// <param name="id"></param>
        public void ReleaseHold(Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CommonHelper commonHelper = new CommonHelper(this.OrganizationService);
            commonHelper.UpdateState(id, this.EntityName, (int)ContractStateCode.Invoiced, (int)ContractInvoicedStatusCode.Invoiced);
        }

        /// <summary>
        /// Copy a <c>Contract</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.clonecontractrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id">
        /// <c>Contract</c> Id to be copied
        /// </param>
        /// <param name="includeCanceledLines">
        /// Indicates whether the canceled line items of the originating contract are to be included in the copy (clone) contract.
        /// Default value is <c>false</c>.
        /// </param>
        /// <returns>
        /// Returns cloned (created) <c>Contract</c> in <see cref="CloneContractResponse.Entity"/> property.
        /// </returns>
        public CloneContractResponse Copy(Guid id, bool includeCanceledLines = false)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            CloneContractRequest request = new CloneContractRequest()
            {
                ContractId = id,
                IncludeCanceledLines = includeCanceledLines
            };

            return (CloneContractResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Cancel a <c>Contract</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.cancelcontractrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id"><c>Conract</c> Id</param>
        /// <param name="cancellationDate">
        /// Contract cancellation date. 
        /// You can pass <c>null</c> value, default value is <see cref="DateTime.UtcNow"/> or <see cref="DateTime.Now"/> depend on <c>useUtc</c> parameter on constructor.
        /// </param>
        /// <param name="status">
        /// <see cref="ContractCanceledStatusCode"/> status code
        /// </param>
        /// <param name="customStatusCode">If you're using your custom statuscodes set this, otherwise you can set "0 (zero)" or null</param>
        /// <returns>
        /// <see cref="CancelContractResponse"/>
        /// </returns>
        public CancelContractResponse Cancel(Guid id, DateTime? cancellationDate, ContractCanceledStatusCode status, int customStatusCode = 0)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            DateTime now = this._useUtc ? DateTime.UtcNow : DateTime.Now;
            int statusCode = (int)status;

            if (status == ContractCanceledStatusCode.CustomStatusCode)
            {
                ExceptionThrow.IfNegative(customStatusCode, "customStatusCode");
                statusCode = customStatusCode;
            }

            CancelContractRequest request = new CancelContractRequest()
            {
                ContractId = id,
                Status = new OptionSetValue(statusCode),
                CancelDate = cancellationDate.HasValue ? cancellationDate.Value : now
            };

            return (CancelContractResponse)this.OrganizationService.Execute(request);
        }

        /// <summary>
        /// Renew a <c>Contract</c> and create the <c>Contract Details</c> for a new contract.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.renewcontractrequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="id">
        /// <c>Contract</c> Id to be renewed
        /// </param>
        /// <param name="includeCanceledLines">
        /// Indicates whether the canceled line items of the original contract should be included in the renewed contract.
        /// Default value is <c>false</c>.
        /// </param>
        /// <returns>
        /// Returns renewed (created) <c>Contract</c> in <see cref="RenewContractResponse.Entity"/> property.
        /// </returns>
        public RenewContractResponse Renew(Guid id, bool includeCanceledLines = false)
        {
            ExceptionThrow.IfGuidEmpty(id, "id");

            RenewContractRequest request = new RenewContractRequest()
            {
                ContractId = id,
                IncludeCanceledLines = includeCanceledLines,
                Status = (int)ContractDraftStatusCode.Draft
            };

            return (RenewContractResponse)this.OrganizationService.Execute(request);
        }

        #endregion
    }
}
