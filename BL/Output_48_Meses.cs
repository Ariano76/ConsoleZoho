using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace BL
{
    class Output_48_Meses
    {
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");


        IDataReader Leer_Ultimos_48_Meses(string Cab, string xPeriodos)
        {
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("")
            return
        }
    }
}
