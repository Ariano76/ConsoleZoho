using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;

using Microsoft.Practices.EnterpriseLibrary.Data;

namespace BL
{
    public class BD_Zoho
    {
        public int[] sPeriodoActual = new int[2];
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor ;
        public readonly string[] sPeriodos48Meses = new string[2];
        public readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_Region_NSE = new double[10, 48];
        public double[,] sdata48Meses_x_Capital_Provincia = new double[2, 48];
        public double[,] sdata48Meses_x_Total = new double[1, 48];

        private readonly DateTime[] Periodos = new DateTime[7];
        string sSource = "Dashboard ZOHO";
        string sLog = "ZOHO";
        string sEvent = "Mensaje de nuestro evento";
        int xMesFin;
        public string resultadoBD;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        static void Main()
        {            
            
        }

        static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file.
            return "Data Source=PEDT0108243; Initial Catalog=BIP; Integrated Security=SSPI";
        }

        public void Periodo_Actual(int pAño, int pMes)
        {
            //'RECUPERA EL ID DEL PERIODO CORRESPONDIENTE AL AÑO Y MES ANTERIOR
            
            if (!EventLog.SourceExists(sSource))
            {
                EventLog.CreateEventSource(sSource, sLog);
            }

            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                string Sql;
                Sql = "EXEC _SP_ID_PERIODO_AÑO_ACTUAL " + pAño + "," + pMes;

                //SqlDataAdapter adapter = new SqlDataAdapter();
                //adapter.TableMappings.Add("Table", "Periodos");

                SqlCommand cmd = new SqlCommand("_SP_ID_PERIODO_AÑO_ACTUAL", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@Año", SqlDbType.Int);
                cmd.Parameters["@Año"].Value = pAño;

                cmd.Parameters.Add("@Mes", SqlDbType.Int);
                cmd.Parameters["@Mes"].Value = pMes;
                try
                {
                    con.Open();
                    sPeriodoActual[0] = (Int32)cmd.ExecuteScalar();
                    xMesFin = (Int32)cmd.ExecuteScalar();
                    //sPeriodoActual[0] = 66;
                }
                catch (Exception ex)
                {
                    EventLog.WriteEntry(sSource, ex.Message, EventLogEntryType.Error,666);
                }                               
            }
        }

        public void Calcular_Periodos(int pAño, int pMes)
        {
            int xPeriodo;
            double Cosmeticos;
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xPeriodo = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT SUM(FACTOR_UNIDAD_VALORIZADO) MONEDA FROM BASE_REGIONES WHERE IDPERIODO >= " + xPeriodo + " AND IDMONEDA = 1 ");
            Cosmeticos = Double.Parse(db.ExecuteScalar(cmd).ToString());

            cmd = db.GetSqlStringCommand("SELECT IDMARCA, MARCA, RANKING, SUM(FACTOR_UNIDAD_VALORIZADO) MONEDA FROM BASE_REGIONES " +
                "WHERE IDPERIODO >= " + xPeriodo + "AND IDMONEDA = 1 GROUP BY IDMARCA, MARCA, RANKING ORDER BY RANKING ASC, MONEDA DESC");

            using (IDataReader reader = db.ExecuteReader(cmd))
            {                
                DataTable dt = new DataTable();
                dt.Load(reader);
                sShareValor = new double[dt.Rows.Count];
                int index = 0;
                foreach (DataRow item in dt.Rows)
                {
                    sShareValor[index] = Double.Parse(item[3].ToString()) / Cosmeticos * 100; 
                    index++;
                }
                //while (reader.Read())
                //{
                //    sShareValor[index] = Double.Parse(reader[3].ToString()) / Cosmeticos * 100;
                //    index++;
                //}
            }
        }

        public DateTime[] Restar_Meses_Fechas(int Anno, int Mes)
        {
            System.DateTime FechaProceso = new System.DateTime(Anno, Mes, 15);
            System.DateTime date2 = DateTime.Today;

            // diff1 gets 185 days, 14 hours, and 47 minutes.
            System.TimeSpan diff1 = date2.Subtract(FechaProceso);

            // date4 gets 4/9/1996 5:55:00 PM.
            System.DateTime date4 = date2.Subtract(diff1);

            // diff2 gets 55 days 4 hours and 20 minutes.
            System.TimeSpan diff2 = date2 - FechaProceso;

            // date5 gets 4/9/1996 5:55:00 PM.
            System.DateTime date5 = FechaProceso - diff2;

            //restar un año a una fecha

            //DateTime oneYearAgoToday = DateTime.Now.AddYears(-1);
            DateTime oneYearAgoToday = FechaProceso.AddYears(-1);
            DateTime twoYearAgoToday = FechaProceso.AddYears(-2);
            DateTime sixMonthAgoToday = FechaProceso.AddMonths(-5);
            DateTime sixMonthTwoYearAgo = oneYearAgoToday.AddMonths(-5);
            DateTime threeMonthAgoToday = FechaProceso.AddMonths(-2);
            DateTime threeMonthTwoYearAgo = oneYearAgoToday.AddMonths(-2);
            DateTime fourYearAgo = FechaProceso.AddYears(-4);
            //Subtracting a week:
            DateTime weekago = DateTime.Now.AddDays(-7);
            Periodos[0] = oneYearAgoToday;
            Periodos[1] = twoYearAgoToday;
            Periodos[2] = sixMonthAgoToday;
            Periodos[3] = sixMonthTwoYearAgo;
            Periodos[4] = threeMonthAgoToday;
            Periodos[5] = threeMonthTwoYearAgo;
            Periodos[6] = fourYearAgo.AddMonths(1);

            return Periodos;
        }

        public string[] Obtener_Ultimos_48_meses(int pAño, int pMes)
        {
            int xMesInicial;
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin );
            
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice=0;
                while (reader.Read())
                {
                    xPeriodos.Append("["+reader[0].ToString()+"],");
                    sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos48Meses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos48Meses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos48Meses;
        }


        public void Leer_Ultimos_48_Meses(string Cab, string xPeriodos)
        {            
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES (CAPITAL Y PROVINCIA) */

            string consulta = @"SELECT REGION, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION";

            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    int cols = reader.FieldCount;
                    int rows = 0;
                    resultadoBD = "No se insertaron los datos";

                    while (reader.Read())
                    {
                        Debug.Write(reader[0].ToString());
                        for (int i = 1; i < cols; i++)
                        {
                            sdata48Meses_x_Region_NSE[rows, i] = double.Parse(reader[i].ToString());
                            Debug.WriteLine(sCabecera48Meses[i-1]);

                            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                            {
                                db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, double.Parse(reader[i].ToString()));
                                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i - 1]);
                                //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                                //db_Zoho.ExecuteNonQuery(cmd_1);
                                db_Zoho.ExecuteNonQuery(cmd_1);
                            }
                        }
                        rows++;
                    }
                    for (int i = 1; i < cols; i++)
                    {
                        Debug.WriteLine(sdata48Meses_x_Region_NSE[0, i]);
                        sdata48Meses_x_Region_NSE[2, i] = sdata48Meses_x_Region_NSE[0, i] + sdata48Meses_x_Region_NSE[1, i];
                        using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                        {
                            db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Region_NSE[0, i] + sdata48Meses_x_Region_NSE[1, i]);
                            db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i-1]);
                            //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                            //db_Zoho.ExecuteNonQuery(cmd_1);
                            db_Zoho.ExecuteNonQuery(cmd_1);
                        }
                    }
                }

            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double capital = 0;
            double provinci = 0;

            string consulta = @"SELECT REGION, NSE, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION,NSE";

            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    int cols = reader.FieldCount;
                    int rows = 0;
                    double valor;
                    resultadoBD = "No se insertaron los datos";

                    while (reader.Read())
                    {
                        //Debug.Write(reader[0].ToString());
                        /* INICIA EN 2 PARA NO LEER COLUMNA REGION Y NSE*/
                        for (int i = 2; i < cols; i++)
                        {                            
                            if (reader[i] == DBNull.Value)
                            {
                                valor = 0;
                            }
                            else
                            {
                                valor = double.Parse(reader[i].ToString());
                            }

                            sdata48Meses_x_Region_NSE[rows, i - 2] = valor;

                            if (reader[0].ToString() == "CAPITAL")
                            {                                
                                sdata48Meses_x_Capital_Provincia[0, i - 2] = sdata48Meses_x_Capital_Provincia[0, i - 2] + valor;
                            }
                            else
                            {
                                sdata48Meses_x_Capital_Provincia[1, i - 2] = sdata48Meses_x_Capital_Provincia[1, i - 2] + valor;
                            }

                            //Debug.WriteLine(sCabecera48Meses[i - 1]);
                            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                            {
                                //db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, double.Parse(reader[i].ToString()));
                                db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, valor);
                                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i-2]);
                                //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                                db_Zoho.ExecuteNonQuery(cmd_1);
                            }
                        }
                        rows++;
                    }

                    /* INSERTANDO VALORES DEL TOTAL CAPITAL A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        //Debug.WriteLine(sdata48Meses_x_Capital_Provincia[0, i] + " - " + sdata48Meses_x_Capital_Provincia[1, i]);   
                        //OBTIENE TOTAL PAIS Y LO ALMACENA EN ARRAY sdata48Meses_x_Total
                        sdata48Meses_x_Total[0, i] = sdata48Meses_x_Capital_Provincia[0, i] + sdata48Meses_x_Capital_Provincia[1, i]; 
                        using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                        {
                            db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Capital_Provincia[0, i]);
                            db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i]);
                            //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                            db_Zoho.ExecuteNonQuery(cmd_1);
                        }
                    }
                    
                    /* INSERTANDO VALORES DEL TOTAL PROVINCIA A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                        {
                            db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Capital_Provincia[1, i]);
                            db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i]);
                            db_Zoho.ExecuteNonQuery(cmd_1);
                        }
                    }

                    for (int i = 0; i < sdata48Meses_x_Total.GetLength(1); i++)
                    {
                        //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                        /* INSERTANDO VALORES DEL TOTAL PAIS A BD ZOHO*/
                        using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                        {
                            db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Total[0, i]);
                            db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i]);
                            //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                            db_Zoho.ExecuteNonQuery(cmd_1);
                        }
                    }

                    for (int id = 0; id < 5; id++)
                    {
                        for (int i = 0; i < sdata48Meses_x_Region_NSE.GetLength(1); i++)
                        {
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
                            {
                                db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Region_NSE[id, i] + sdata48Meses_x_Region_NSE[id + 5, i]);
                                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i]);
                                //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
                                db_Zoho.ExecuteNonQuery(cmd_1);
                            }
                        }
                    }


                }

            }
        }

    }
}
