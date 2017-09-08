using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Xrm.Sdk;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;

namespace XrmLibrary.EntityHelpers.Activity
{
    /// <summary>
    /// Fluent object to easy-create MS CRM <c>Email</c> activity.
    /// </summary>
    public class XrmEmail
    {
        #region | Private Definitions |

        bool _isInit = false;
        bool _directionCode;
        int _priority;

        KeyValuePair<string, Guid> _from;
        List<Tuple<ToEntityType, Guid, string>> _to;
        List<Tuple<ToEntityType, Guid, string>> _cc;
        List<Tuple<ToEntityType, Guid, string>> _bcc;
        string _subject;
        string _body;
        KeyValuePair<string, Guid> _regardingObject;
        Dictionary<String, object> _customAttributeList;

        #endregion

        #region | Enums |

        /// <summary>
        /// <c>Email</c> activity 's default <c>To</c> party values
        /// </summary>
        public enum ToEntityType
        {
            [Description("undefined")]
            Undefined,

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

        #endregion

        #region | Constructors |

        /// <summary>
        /// 
        /// </summary>
        public XrmEmail()
        {
            _isInit = true;
            _directionCode = Convert.ToBoolean(DirectionCode.Outgoing.Description());
            _priority = (int)PriorityCode.Normal;
            _to = new List<Tuple<ToEntityType, Guid, string>>();
            _cc = new List<Tuple<ToEntityType, Guid, string>>();
            _bcc = new List<Tuple<ToEntityType, Guid, string>>();
            _customAttributeList = new Dictionary<string, object>();
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add <c>From</c>.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail From(FromEntityType entityType, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "From.Id");

            if (string.IsNullOrEmpty(_from.Key) && _from.Value.IsGuidEmpty())
            {
                _from = new KeyValuePair<string, Guid>(entityType.Description(), id);
            }
            else
            {
                throw new ArgumentException("You can only add one record FROM data. Please NULL to list before add new record");
            }

            return this;
        }

        /// <summary>
        /// Add <c>To</c>.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail To(ToEntityType entityType, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "To.Id");

            _to.Add(new Tuple<ToEntityType, Guid, string>(entityType, id, string.Empty));

            return this;
        }

        /// <summary>
        /// Add <c>To</c> with bulk Id list.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="idList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail To(ToEntityType entityType, List<Guid> idList)
        {
            ExceptionThrow.IfNull(idList, "To.Id List");

            foreach (var item in idList)
            {
                _to.Add(new Tuple<ToEntityType, Guid, string>(entityType, item, string.Empty));
            }

            return this;
        }

        /// <summary>
        /// Add <c>To</c> with email address.
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns><<see cref="XrmEmail"/>/returns>
        public XrmEmail To(string emailAddress)
        {
            ExceptionThrow.IfNullOrEmpty(emailAddress, "To EmailAddress");

            _to.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, emailAddress));

            return this;
        }

        /// <summary>
        /// Add <c>To</c> with bulk email address list.
        /// </summary>
        /// <param name="emailAddressList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail To(List<string> emailAddressList)
        {
            ExceptionThrow.IfNull(emailAddressList, "To Email Address List");

            foreach (string item in emailAddressList)
            {
                _to.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, item));
            }

            return this;
        }

        /// <summary>
        /// Add <c>CC</c>.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Cc(ToEntityType entityType, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "Cc.Id");

            _cc.Add(new Tuple<ToEntityType, Guid, string>(entityType, id, string.Empty));

            return this;
        }

        /// <summary>
        /// Add <c>CC</c> with bulk Id list.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="idList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Cc(ToEntityType entityType, List<Guid> idList)
        {
            ExceptionThrow.IfNull(idList, "Cc.Id List");

            foreach (var item in idList)
            {
                _cc.Add(new Tuple<ToEntityType, Guid, string>(entityType, item, string.Empty));
            }

            return this;
        }

        /// <summary>
        /// Add <c>CC</c> with email address.
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Cc(string emailAddress)
        {
            ExceptionThrow.IfNullOrEmpty(emailAddress, "CC EmailAddress");

            _cc.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, emailAddress));

            return this;
        }

        /// <summary>
        /// Add <c>CC</c> with bulk email address list.
        /// </summary>
        /// <param name="emailAddressList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Cc(List<string> emailAddressList)
        {
            ExceptionThrow.IfNull(emailAddressList, "CC Email Address List");

            foreach (string item in emailAddressList)
            {
                _cc.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, item));
            }

            return this;
        }

        /// <summary>
        /// Add <c>BCC</c>.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Bcc(ToEntityType entityType, Guid id)
        {
            ExceptionThrow.IfGuidEmpty(id, "Bcc.Id");

            _bcc.Add(new Tuple<ToEntityType, Guid, string>(entityType, id, string.Empty));

            return this;
        }

        /// <summary>
        /// Add <c>BCC</c> with bulk Id list.
        /// </summary>
        /// <param name="entityType"></param>
        /// <param name="idList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Bcc(ToEntityType entityType, List<Guid> idList)
        {
            ExceptionThrow.IfNull(idList, "Bcc.Id List");

            foreach (var item in idList)
            {
                _bcc.Add(new Tuple<ToEntityType, Guid, string>(entityType, item, string.Empty));
            }

            return this;
        }

        /// <summary>
        /// Add <c>BCC</c> with email address.
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Bcc(string emailAddress)
        {
            ExceptionThrow.IfNullOrEmpty(emailAddress, "BCC EmailAddress");

            _bcc.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, emailAddress));

            return this;
        }

        /// <summary>
        /// Add <c>BCC</c> with bulk email address list.
        /// </summary>
        /// <param name="emailAddressList"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Bcc(List<string> emailAddressList)
        {
            ExceptionThrow.IfNull(emailAddressList, "BCC Email Address List");

            foreach (string item in emailAddressList)
            {
                _bcc.Add(new Tuple<ToEntityType, Guid, string>(ToEntityType.Undefined, Guid.Empty, item));
            }

            return this;
        }

        /// <summary>
        /// Add <c>Subject</c>.
        /// </summary>
        /// <param name="subject"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Subject(string subject)
        {
            ExceptionThrow.IfNullOrEmpty(subject, "subject");

            _subject = subject.Trim();

            return this;
        }

        /// <summary>
        /// Add <c>Body</c>.
        /// </summary>
        /// <param name="body"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Body(string body)
        {
            ExceptionThrow.IfNullOrEmpty(body, "body");

            _body = body.Trim();

            return this;
        }

        /// <summary>
        /// Add <c>Regarding</c>.
        /// </summary>
        /// <param name="entityLogicalName"></param>
        /// <param name="id"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Regarding(string entityLogicalName, Guid id)
        {
            ExceptionThrow.IfNullOrEmpty(entityLogicalName, "Regarding.EntityLogicalName");
            ExceptionThrow.IfGuidEmpty(id, "Regarding.Id");

            if (string.IsNullOrEmpty(_regardingObject.Key) && _regardingObject.Value.IsGuidEmpty())
            {
                _regardingObject = new KeyValuePair<string, Guid>(entityLogicalName.ToLower().Trim(), id);
            }

            return this;
        }

        /// <summary>
        /// Add <c>Direction</c>.
        /// </summary>
        /// <param name="direction"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Direction(DirectionCode direction)
        {
            _directionCode = Convert.ToBoolean(direction.Description());
            return this;
        }

        /// <summary>
        /// Add <c>Priority</c>.
        /// </summary>
        /// <param name="priority"></param>
        /// <returns><see cref="XrmEmail"/></returns>
        public XrmEmail Priority(PriorityCode priority)
        {
            _priority = (int)priority;
            return this;
        }

        /// <summary>
        /// Add additional attribute
        /// </summary>
        /// <param name="attributeName">Attribute name</param>
        /// <param name="value">
        /// Attribute value by MS CRM datatypes (<see cref="String"/>, <see cref="int"/> or <see cref="EntityReference"/>, <see cref="OptionSetValue"/>)
        /// </param>
        /// <returns>
        /// <see cref="XrmEmail"/>
        /// </returns>
        public XrmEmail AddCustomAttribute(string attributeName, object value)
        {
            ExceptionThrow.IfNullOrEmpty(attributeName, "attributeName");

            _customAttributeList.Add(attributeName, value);

            return this;
        }

        /// <summary>
        /// Convert <see cref="XrmEmail"/> to <see cref="Entity"/> for <c>Email Activity</c>.
        /// </summary>
        /// <returns><see cref="Entity"/></returns>
        public Entity ToEntity()
        {
            Entity result = null;

            if (!_isInit)
            {
                throw new InvalidOperationException("Please INIT (call constructor of XrmEmail) before call ToEntity method.");
            }

            ExceptionThrow.IfNullOrEmpty(_from.Key, "From.EntityType");
            ExceptionThrow.IfGuidEmpty(_from.Value, "From.Id");
            ExceptionThrow.IfNull(_to, "To");
            ExceptionThrow.IfNullOrEmpty(_to.ToArray(), "To");
            ExceptionThrow.IfNullOrEmpty(_subject, "Subject");
            ExceptionThrow.IfNullOrEmpty(_body, "Body");

            result = new Entity("email");

            Entity from = new Entity("activityparty");
            from["partyid"] = new EntityReference(_from.Key, _from.Value);

            result["subject"] = _subject;
            result["description"] = _body;
            result["from"] = new[] { from };
            result["to"] = CreateActivityParty(_to);
            result["cc"] = CreateActivityParty(_cc);
            result["bcc"] = CreateActivityParty(_bcc);
            result["directioncode"] = _directionCode;
            result["prioritycode"] = new OptionSetValue(_priority);
            result["statuscode"] = new OptionSetValue(1);
            //INFO : In Microsoft Dynamics 365(online & on - premises), the Email.StatusCode attribute cannot be null.

            if (!string.IsNullOrEmpty(_regardingObject.Key) && !_regardingObject.Value.IsGuidEmpty())
            {
                result["regardingobjectid"] = new EntityReference(_regardingObject.Key, _regardingObject.Value);
            }

            if (_customAttributeList != null && _customAttributeList.Keys.Count > 0)
            {
                foreach (KeyValuePair<string, object> item in _customAttributeList)
                {
                    result[item.Key] = item.Value;
                }
            }

            return result;
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Create <c>activityparty</c>.
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        EntityCollection CreateActivityParty(List<Tuple<ToEntityType, Guid, string>> list)
        {
            EntityCollection result = null;

            if (!list.IsNullOrEmpty())
            {
                result = new EntityCollection();

                foreach (var item in list)
                {
                    Entity p = new Entity("activityparty");

                    if (!item.Item2.IsGuidEmpty() && item.Item1 != ToEntityType.Undefined)
                    {
                        p["partyid"] = new EntityReference(item.Item1.Description(), item.Item2);
                    }
                    else if (!string.IsNullOrEmpty(item.Item3))
                    {
                        p["addressused"] = item.Item3.Trim();
                    }

                    result.Entities.Add(p);
                }
            }

            return result;
        }

        #endregion
    }
}
