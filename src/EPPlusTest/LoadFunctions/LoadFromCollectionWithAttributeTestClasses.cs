using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    [EpplusTable]
    public abstract class OrganizationBase
    {
        [EpplusTableColumn(Header = "Org Level 3", Order = 1)]
        public virtual string OrgLevel3 { get; set; }

        [EpplusTableColumn(Header = "Org Level 4", Order = 2)]
        public string OrgLevel4 { get; set; }

        [EpplusTableColumn(Header = "Org Level 5", Order = 3)]
        public string OrgLevel5 { get; set; }
    }

    [EpplusTable]
    public class Organization
    {
        [EpplusTableColumn(Header = "Org Level 3", Order = 1)]
        public string OrgLevel3 { get; set; }

        [EpplusTableColumn(Header = "Org Level 4", Order = 2)]
        public string OrgLevel4 { get; set; }

        [EpplusTableColumn(Header = "Org Level 5", Order = 3)]
        public string OrgLevel5 { get; set; }
    }

    [EpplusTable]
    public class OrganizationReversedSortOrder
    {
        [EpplusTableColumn(Header = "Org Level 3", Order = 3)]
        public string OrgLevel3 { get; set; }

        [EpplusTableColumn(Header = "Org Level 4", Order = 2)]
        public string OrgLevel4 { get; set; }

        [EpplusTableColumn(Header = "Org Level 5", Order = 1)]
        public string OrgLevel5 { get; set; }
    }

    [EpplusTable]
    public class OrganizationSubclass : OrganizationBase
    {
        public override string OrgLevel3 { get; set; }
    }

    [EpplusTable]
    public class Outer
    {
        [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 1)]
        public DateTime? ApprovedUtc { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public Organization Organization { get; set; }

        [EpplusTableColumn(Header = "Acknowledged...", Order = 3)]
        public bool Acknowledged { get; set; }
    }

    [EpplusTable(PrintHeaders = true)]
    public class OuterWithHeaders
    {
        [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 1)]
        public DateTime? ApprovedUtc { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public Organization Organization { get; set; }

        [EpplusTableColumn(Header = "Acknowledged...", Order = 3)]
        public bool Acknowledged { get; set; }
    }

    [EpplusTable(PrintHeaders = true)]
    [EPPlusTableColumnSortorder(Properties = new string[]
    {
        nameof(Acknowledged),
        "Organization.OrgLevel4",
        nameof(ApprovedUtc)
    })]
    public class OuterWithSortOrderOnClassLevelV1
    {
        [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 1)]
        public DateTime? ApprovedUtc { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public Organization Organization { get; set; }

        [EpplusTableColumn(Header = "Acknowledged...", Order = 3)]
        public bool Acknowledged { get; set; }
    }

    [EpplusTable]
    public class OuterReversedSortOrder
    {
        [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 3)]
        public DateTime? ApprovedUtc { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public OrganizationReversedSortOrder Organization { get; set; }

        [EpplusTableColumn(Header = "Acknowledged...", Order = 1)]
        public bool Acknowledged { get; set; }
    }

    [EpplusTable]
    public class OuterSubclass
    {
        [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 3)]
        public DateTime? ApprovedUtc { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public OrganizationSubclass Organization { get; set; }

        [EpplusTableColumn(Header = "Acknowledged...", Order = 1)]
        public bool Acknowledged { get; set; }
    }
}
