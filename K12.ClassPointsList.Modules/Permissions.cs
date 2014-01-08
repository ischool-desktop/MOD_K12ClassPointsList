using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12ClassPointsList.Modules
{
    class Permissions
    {
        public static string 班級點名單_套表列印 { get { return "K12.ClassPointsList.Modules.ClassPointsList.cs"; } }

        public static bool 班級點名單_套表列印權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級點名單_套表列印].Executable;
            }
        }
    }
}
