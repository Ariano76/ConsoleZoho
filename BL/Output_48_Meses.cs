using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Diagnostics;

namespace BL
{
    public class Output_48_Meses
    {
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        public double[,] sdata48Meses_x_Region = new double[2,49];
   
        public void Leer_Ultimos_48_Meses(string Cab, string xPeriodos)
        {
            //string consulta = @"SELECT REGION, @Cab "+
            //    "FROM ( SELECT PERIODO, REGION, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
            //    "WHERE IDPERIODO IN (@xPeriodos) AND IDMONEDA = 2 ) AS SourceTable "+
            //    "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) "+
            //    "FOR PERIODO IN (@Cab)) AS pvt "+
            //    "ORDER BY REGION";
            string consulta = @"SELECT REGION, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION";
            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                //db.AddInParameter(cmd, "@Cab", DbType.String, Cab);
                //db.AddInParameter(cmd, "@xPeriodos", DbType.String, xPeriodos);

                using (IDataReader reader = db.ExecuteReader(cmd)) 
                {
                    int cols = reader.FieldCount;
                    while (reader.Read())
                    {
                        Debug.Write(reader[0].ToString());
                        for (int i = 1; i < cols; i++)
                        {
                            sdata48Meses_x_Region[0, i] = double.Parse(reader[i].ToString());
                        }
                    }

                }
            }
        }
    }
}
