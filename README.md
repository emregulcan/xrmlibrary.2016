# xrmlibrary.2016
Microsoft Dynamics CRM 2016 / Dynamics 365 helper library

This project includes 2 helper library for Microsoft Dynamics CRM 2016 / Dynamics 365.
* **XrmLibrary.EntityHelpers** : Microsoft Dynamics CRM 2016 Entity Helper. This library includes almost full MS CRM 2016 SDK methods and UI methods.
* **XrmLibrary.ExtensionMethods** : Microsoft Dynamics CRM 2016 SDK extension methods

For more information please visit https://github.com/emregulcan/xrmlibrary.2016/wiki

Please read important notices about [SDK Deprecated Messages](https://github.com/emregulcan/xrmlibrary.2016/wiki/06-SDK-Deprecated-Messages)

Your comments and pull requests are very important and always welcome


# Commit / Release Notes

## v1.4
- [AddCustomAttribute](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Activity/XrmEmail.cs#L382-L389) method added to XrmEmail
- [AddCustomAttribute](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Schedule/Data/XrmAppointment.cs#L266-L273) method added to XrmAppointment

## v1.3
- Strong-named (signed) assembly support added. Please read [this issue](https://github.com/emregulcan/xrmlibrary.2016/issues/3)
- BUG-FIXING
   - Schedule.Data.XrmAppointment
      - "[location](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Schedule/Data/XrmAppointment.cs#L288)" attribute added to converting MS Dynamics entity

## v1.2
- [GetSharedPrincipalsAndAccess](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Common/CommonHelper.cs#L524-L544) method added to CommonHelper.

- [ReassignOwnership](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Common/CommonHelper.cs#L559-L571) method added to CommonHelper.

- [Add](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Sales/InvoiceProductHelper.cs#L45-L57) method added to InvoiceProductHelper to retrieve products from Opportunity.

## v1.1
- [GetMemberList](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Marketing/ListHelper.cs#L281-L310) method added to Marketing.ListHelper. This retrieves all members in marketing list (without regarding Static or Dynamic) and returns ListMemberResult class. 
- [BaseEntityHelper.Get(string fetchxml)](https://github.com/emregulcan/xrmlibrary.2016/blob/master/XrmLibrary.EntityHelpers/Common/BaseEntityHelper.cs#L101-L107) mofied to use "FetchExpression".
- [XrmLibrary.ExtensionMethods.2016](https://github.com/emregulcan/xrmlibrary.2016/tree/master/XrmLibrary.ExtensionMethods) assembly dependecy removed.

## v1.0
- First release
