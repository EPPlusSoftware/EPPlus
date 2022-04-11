using OfficeOpenXml.Attributes;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    [
    EpplusTable(AutofitColumns = true, PrintHeaders = true, TableStyle = TableStyles.Medium2),
    EPPlusTableColumnSortOrder(Properties = new string[] {
        nameof(PlatformName), nameof(PchDieName), nameof(OtherDieName), nameof(Stepping), nameof(MilestoneDay),
        nameof(CollateralOwner), nameof(MissionControlLead), nameof(CreatedUtc), nameof(ModifiedUtc)
    })
]
    public class IntegratedPlatformExcelRow
    {
        [EpplusNestedTableColumn(HeaderPrefix = "Collateral Owner")]
        public WorkerDTO CollateralOwner { get; set; }

        [EpplusTableColumn(Header = "Created (GMT)", NumberFormat = "yyyy-MM-dd HH:MM")]
        public DateTime CreatedUtc { get; set; }

        [EpplusTableColumn(Header = "Milestone Day")]
        public string MilestoneDay { get; set; }

        [EpplusNestedTableColumn (HeaderPrefix = "Mission Control Lead")]
        public WorkerDTO MissionControlLead { get; set; }

        [EpplusTableColumn(Header = "Modified (GMT)", NumberFormat = "yyyy-MM-dd HH:MM")]
        public DateTime ModifiedUtc { get; set; }

        [EpplusTableColumn(Header = "SOC/CPU Die Name")]
        public string OtherDieName { get; set; }

        [EpplusTableColumn(Header = "PCH Die Name")]
        public string PchDieName { get; set; }

        [EpplusTableColumn(Header = "Product Family")]
        public string PlatformName { get; set; }

        public string Stepping { get; set; }
    }

    [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Name), nameof(Email), nameof(WWID) })]
    public sealed class WorkerDTO
    {
        [EpplusIgnore]
        public bool Active { get; set; }

        public string Email { get; set; }

        public string Name { get; set; }

        public int WWID { get; set; }
    }

    public static class ExcelItems {
     
        public static IEnumerable<IntegratedPlatformExcelRow> GetItems1()
        {
            return new List<IntegratedPlatformExcelRow>
            {
                new IntegratedPlatformExcelRow
                {
                    CollateralOwner = new WorkerDTO
                    {
                        Email = "col@owner.com",
                        Name = "Collateral Owner"
                    },
                    CreatedUtc = DateTime.Now,
                    MilestoneDay = "Mile Stone Day 1",
                    MissionControlLead = new WorkerDTO
                    {
                        Email = "miss@control.com",
                        Name = "Misson Control Lead"
                    },
                    ModifiedUtc = DateTime.Now,
                    OtherDieName = "Other Die name 1",
                    PchDieName = "Pch Die Name 1",
                    PlatformName = "Platform name 1",
                    Stepping = "Stepping 1"
                }
            };
        }
    }
}
