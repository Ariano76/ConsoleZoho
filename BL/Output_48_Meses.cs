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
        Database db_Zoho = factory.Create("ZOHO");
        public double[,] sdata48Meses_x_Region = new double[3,49];
        //public string resultadoBD;
        BD_Zoho obj48 = new BD_Zoho();

        //public void Leer_Ultimos_48_Meses(string Cab, string xPeriodos)
        //{
        //    //string consulta = @"SELECT REGION, @Cab "+
        //    //    "FROM ( SELECT PERIODO, REGION, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
        //    //    "WHERE IDPERIODO IN (@xPeriodos) AND IDMONEDA = 2 ) AS SourceTable "+
        //    //    "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) "+
        //    //    "FOR PERIODO IN (@Cab)) AS pvt "+
        //    //    "ORDER BY REGION";
        //    string consulta = @"SELECT REGION, " + @Cab +
        //        " FROM ( SELECT PERIODO, REGION, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
        //        "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
        //        "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
        //        "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
        //        "ORDER BY REGION";
                     
        //    using (DbCommand cmd = db.GetSqlStringCommand(consulta))
        //    {
        //        using (IDataReader reader = db.ExecuteReader(cmd)) 
        //        {
        //            int cols = reader.FieldCount;
        //            int rows = 0;
        //            resultadoBD = "No se insertaron los datos";

        //            while (reader.Read())
        //            {
        //                Debug.Write(reader[0].ToString());
        //                for (int i = 1; i < cols; i++)
        //                {
        //                    sdata48Meses_x_Region[rows, i] = double.Parse(reader[i].ToString());
        //                    Debug.WriteLine(obj48.sCabecera48Meses[i]);

        //                    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA") )
        //                    {
        //                        db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, double.Parse(reader[i].ToString()));
        //                        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, obj48.sCabecera48Meses[i-1]);
        //                        //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
        //                        //db_Zoho.ExecuteNonQuery(cmd_1);
        //                        db_Zoho.ExecuteNonQuery(cmd_1);
        //                    }                                                       
        //                }
        //                rows++;
        //            }
        //            for (int i = 0; i < cols; i++)
        //            {
        //                sdata48Meses_x_Region[2, i] = sdata48Meses_x_Region[0, i] + sdata48Meses_x_Region[1, i];
        //            }
        //        }
                
        //    }
        //}
    }
}
