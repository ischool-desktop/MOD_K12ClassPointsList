using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;

namespace K12ClassPointsList.Modules
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {

            RibbonBarItem bh = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            bh["報表"]["學務相關報表"]["班級點名單(套表列印)"].Enable = Permissions.班級點名單_套表列印權限;
            bh["報表"]["學務相關報表"]["班級點名單(套表列印)"].Click += delegate
            {
                ClassPointsList cpl = new ClassPointsList();
                cpl.ShowDialog();
            };

            Catalog detail = RoleAclSource.Instance["班級"]["報表"];
            detail.Add(new ReportFeature(Permissions.班級點名單_套表列印, "班級點名單(套表列印)"));
        }
    }
}
