using System;
using System.Collections.Generic;
using XrmLibrary.EntityHelpers.Common;
using Sasha.ExtensionMethods;
using Microsoft.Xrm.Sdk;

namespace XrmLibrary.EntityHelpers.Marketing
{
    /// <summary>
    /// <c>ListMemberResult</c> contains <see cref="ListMemberItemDetail"/> data for <c>ListMember</c> entity.
    /// </summary>
    public class ListMemberResult
    {
        #region | Private Definitions |

        ListMemberTypeCode _memberType = ListMemberTypeCode.Undefined;
        List<ListMemberItemDetail> _data = null;

        #endregion

        #region | Properties |

        /// <summary>
        /// Item count
        /// </summary>
        public int Count { get { return this.Data.IsNullOrEmpty() ? 0 : this.Data.Count; } }

        /// <summary>
        /// Member type name (entity logical name)
        /// </summary>
        public ListMemberTypeCode MemberType { get { return _memberType; } }

        /// <summary>
        /// Member data (Id and Name like <see cref="EntityReference"/>).
        /// </summary>
        public List<ListMemberItemDetail> Data { get { return _data; } }

        #endregion

        #region | Constructors |

        /// <summary>
        /// 
        /// </summary>
        public ListMemberResult()
        {
            this._memberType = ListMemberTypeCode.Undefined;
            this._data = new List<ListMemberItemDetail>();
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Add single list member 
        /// </summary>
        /// <param name="memberType"></param>
        /// <param name="item"></param>
        public void Add(ListMemberTypeCode memberType, ListMemberItemDetail item)
        {
            this._memberType = memberType;
            this._data.Add(item);
        }

        /// <summary>
        /// Add bulk list member
        /// </summary>
        /// <param name="memberType"></param>
        /// <param name="data"></param>
        public void Add(ListMemberTypeCode memberType, List<ListMemberItemDetail> data)
        {
            this._memberType = memberType;
            this._data.AddRange(data);
        }

        #endregion
    }

    /// <summary>
    /// 
    /// </summary>
    public class ListMemberItemDetail
    {
        #region | Private Definitions |

        Guid _id = Guid.Empty;
        string _name = string.Empty;

        #endregion

        #region | Properties |

        public Guid Id { get { return this._id; } }
        public string Name { get { return this._name; } }

        #endregion

        #region | Constructors |

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="name"></param>
        public ListMemberItemDetail(Guid id, string name)
        {
            this._id = id;
            this._name = name;
        }

        #endregion
    }
}
