using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace BL
{
    public class Procesos_Genericos
    {
        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db_Zoho = factory.Create("ZOHO");
        int Total;

        public int Leer_Total_Registros_BD()
        {
            using (DbCommand cmd = db_Zoho.GetStoredProcCommand("_SP_TOTAL_REGISTROS"))
            {
                Total = int.Parse(db_Zoho.ExecuteScalar(cmd).ToString());
            }
            return Total;
        }
    }
}
