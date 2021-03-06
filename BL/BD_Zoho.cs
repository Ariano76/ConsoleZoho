﻿using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;

namespace BL
{
    public class BD_Zoho
    {
        public int[] sPeriodoActual = new int[2];
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor;
        //public static readonly string[] sPeriodosYTDMeses = new string[2];     
        public static readonly string[] s3AñosAnteriores = new string[2];
        public static readonly string[] sPeriodos1Meses = new string[2];
        public static readonly string[] sPeriodos1MesesOneYearAgo = new string[2];
        public static readonly string[] sPeriodos3Meses = new string[2];
        public static readonly string[] sPeriodos3MesesOneYearAgo = new string[2];
        public static readonly string[] sPeriodos6Meses = new string[2];
        public static readonly string[] sPeriodos6MesesOneYearAgo = new string[2];
        public static readonly string[] sPeriodos12Meses = new string[2];
        public static readonly string[] sPeriodos12MesesOneYearAgo = new string[2];
        public static readonly string[] sPeriodosYTDMeses = new string[2];
        public static readonly string[] sPeriodosYTDMesesOneYearAgo = new string[2];
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];
        public static readonly string[] sPeriodos_Año_0 = new string[2];
        public static readonly string[] sPeriodos_Año_1 = new string[2];
        public static readonly string[] sPeriodos_Año_2 = new string[2];
        public static readonly string[] sPeriodos_All_Años = new string[2];

        public static readonly string[] sPeriodos48MesesFactor = new string[2];
        //public static readonly string[] sCabecera48MesesFactorCapital = new string[48];       
        //public static readonly string[] sCabecera48MesesFactorCiudad = new string[48];

        public double[,] sdata48Meses_x_Region_NSE = new double[10, 48];
        public double[,] sdata48Meses_x_Capital_Provincia = new double[2, 48];
        public double[,] sdata48Meses_x_Total = new double[1, 48];

        public double[,] sdata48Meses_x_Region_Categoria_Mes = new double[6, 48];
        public double[,] sdata48Meses_x_Region_Modalidad_Mes = new double[4, 48];
        public double[,] sdata48Meses_x_Region_Tipos_Mes = new double[12, 48];


        private readonly DateTime[] Periodos = new DateTime[17];
        string sSource = "Dashboard ZOHO";
        string NSE_, Ciudad_, Mercado_, Periodo;
        int xMesFin_0, xMesFin, xMesFin_1, xMesFin_3, xMesFin_6, xMesFin_12, xMesFin_YTD;
        int xMesInicio, xMesInicial, xMesInicial_1, xMesInicial_3, xMesInicial_6, xMesInicial_12, xMesInicial_YTD;
        public string resultadoBD;
        public int Numero_Meses_en_YTD;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file.
            return "Data Source=PELT0113401; Initial Catalog=BIP; User Id=Procesamiento; Password=P@ssw0rd";
        }

        public void Periodo_Actual(int pAño, int pMes)
        {
            //'RECUPERA EL ID DEL PERIODO CORRESPONDIENTE AL AÑO Y MES ANTERIOR

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
                    EventLog.WriteEntry(sSource, ex.Message, EventLogEntryType.Error, 666);
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
            System.DateTime Mes_1_Año_Actual = new System.DateTime(Anno, 1, 15);
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
            DateTime oneYearAgoToday = FechaProceso.AddYears(-1); //mes actual un año atras
            DateTime twoYearAgoToday = FechaProceso.AddYears(-2); //mes actual dos años atras
            DateTime threeYearAgoToday = FechaProceso.AddYears(-3);
            DateTime sixMonthAgoToday = FechaProceso.AddMonths(-5);
            DateTime EndSixMonthAgoToday = FechaProceso.AddMonths(-5);
            DateTime StartSixMonthAgoToday = oneYearAgoToday.AddMonths(-5);
            DateTime sixMonthTwoYearAgo = oneYearAgoToday.AddMonths(-5);
            DateTime threeMonthAgoToday = FechaProceso.AddMonths(-3);
            DateTime threeMonthTwoYearAgo = oneYearAgoToday.AddMonths(-3);
            DateTime fourYearAgo = FechaProceso.AddYears(-4);
            DateTime ThreeMonthsAgo = FechaProceso.AddMonths(-2);
            DateTime StartThreeMonthAgoToday = oneYearAgoToday.AddMonths(-2);
            DateTime TwelveMonthsAgo = FechaProceso.AddMonths(-11);
            DateTime TwentyFourMonthsAgo = FechaProceso.AddMonths(-23);
            DateTime TwelveMonthTwoYearAgo = oneYearAgoToday.AddMonths(-11);
            DateTime Mes_1_Año_Anterior = Mes_1_Año_Actual.AddYears(-1);

            //Subtracting a week:
            DateTime weekago = DateTime.Now.AddDays(-7);

            Periodos[0] = oneYearAgoToday;
            Periodos[1] = twoYearAgoToday;
            Periodos[8] = threeYearAgoToday;
            Periodos[2] = sixMonthAgoToday;
            Periodos[3] = sixMonthTwoYearAgo;
            Periodos[4] = threeMonthAgoToday;
            Periodos[5] = threeMonthTwoYearAgo;
            Periodos[6] = fourYearAgo.AddMonths(1);
            Periodos[7] = ThreeMonthsAgo;
            Periodos[9] = TwelveMonthsAgo;
            Periodos[10] = TwelveMonthTwoYearAgo;
            Periodos[11] = Mes_1_Año_Actual;
            Periodos[12] = Mes_1_Año_Anterior;
            Periodos[13] = TwentyFourMonthsAgo;
            Periodos[14] = StartSixMonthAgoToday;
            Periodos[15] = EndSixMonthAgoToday;
            Periodos[16] = StartThreeMonthAgoToday;

            return Periodos;
        }

        public string[] Obtener_3_Ultimos_años(int pAño)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();

            pAño -= 3;
            for (int i = 0; i < 3; i++)
            {
                xPeriodos.Append("[" + (pAño + i) + "],");
                xPeriodosInt.Append((pAño + i) + ",");
            }

            s3AñosAnteriores[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            s3AñosAnteriores[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return s3AñosAnteriores;
        }
        //obtiene los periodos del año actual mas los ultimos 3 años.
        public string[] Obtener_periodos_all_años(int pAñoActual)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + (pAñoActual - 3) + " ORDER BY IDPERIODO ASC");
            xMesInicio = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAñoActual + " ORDER BY IDPERIODO DESC");
            xMesFin_0 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                }
            }

            sPeriodos_All_Años[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos_All_Años[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos_All_Años;
        }
        public string[] Obtener_periodos_año_0(int pAño)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + (pAño) + " ORDER BY IDPERIODO ASC");
            xMesInicio = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " ORDER BY IDPERIODO DESC");
            xMesFin_0 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                }
            }

            sPeriodos_Año_0[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos_Año_0[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos_Año_0;
        }
        public string[] Obtener_periodos_año_1(int pAño)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " ORDER BY IDPERIODO ASC");
            xMesInicio = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " ORDER BY IDPERIODO DESC");
            xMesFin_0 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                }
            }

            sPeriodos_Año_1[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos_Año_1[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos_Año_1;
        }
        public string[] Obtener_periodos_año_2(int pAño)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " ORDER BY IDPERIODO ASC");
            xMesInicio = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " ORDER BY IDPERIODO DESC");
            xMesFin_0 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicio + " AND IDPERIODO <= " + xMesFin_0);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                }
            }

            sPeriodos_Año_2[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos_Año_2[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos_Año_2;
        }

        public string[] Obtener_Ultimos_1_mes(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial_1 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO DESC");
            xMesFin_1 = (Int32)(db.ExecuteScalar(cmd));


            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_1 + " AND IDPERIODO <= " + xMesFin_1);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_1 + " AND IDPERIODO <= " + xMesFin_1);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos1Meses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos1Meses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos1Meses;
        }
        public string[] Obtener_Ultimos_1_mes_One_Year_Ago(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial_1 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO DESC");
            xMesFin_1 = (Int32)(db.ExecuteScalar(cmd));


            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_1 + " AND IDPERIODO <= " + xMesFin_1);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_1 + " AND IDPERIODO <= " + xMesFin_1);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos1MesesOneYearAgo[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos1MesesOneYearAgo[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos1MesesOneYearAgo;
        }
        public string[] Obtener_Ultimos_3_meses(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
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

            sPeriodos3Meses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos3Meses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos3Meses;
        }
        public string[] Obtener_Ultimos_3_meses_One_Year_Ago(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + Periodos[16].Year + " AND MES = " + Periodos[16].Month + " ORDER BY IDPERIODO ASC");
            xMesInicial_3 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO DESC");
            xMesFin_3 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_3 + " AND IDPERIODO <= " + xMesFin_3);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_3 + " AND IDPERIODO <= " + xMesFin_3);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos3MesesOneYearAgo[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos3MesesOneYearAgo[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos3MesesOneYearAgo;
        }
        public string[] Obtener_Ultimos_6_meses(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
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

            sPeriodos6Meses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos6Meses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos6Meses;
        }
        public string[] Obtener_Ultimos_6_meses_One_Year_Ago(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + Periodos[14].Year + " AND MES = " + Periodos[14].Month + " ORDER BY IDPERIODO ASC");
            xMesInicial_6 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO DESC");
            xMesFin_6 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_6 + " AND IDPERIODO <= " + xMesFin_6);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_6 + " AND IDPERIODO <= " + xMesFin_6);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos6MesesOneYearAgo[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos6MesesOneYearAgo[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos6MesesOneYearAgo;
        }
        public string[] Obtener_Ultimos_12_meses(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
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

            sPeriodos12Meses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos12Meses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos12Meses;
        }
        public string[] Obtener_Ultimos_12_meses_One_Year_Ago(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + Periodos[13].Year + " AND MES = " + Periodos[13].Month + " ORDER BY IDPERIODO ASC");
            xMesInicial_12 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO DESC");
            xMesFin_12 = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_12 + " AND IDPERIODO <= " + xMesFin_12);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_12 + " AND IDPERIODO <= " + xMesFin_12);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodos12MesesOneYearAgo[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodos12MesesOneYearAgo[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodos12MesesOneYearAgo;
        }
        //public int Numero_Meses_YTD(int pAño, int pMes)
        //{            
        //    int numeroMeses;
        //    DbCommand cmd;
        //    cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
        //    xMesInicial = (Int32)(db.ExecuteScalar(cmd));

        //    cmd = db.GetSqlStringCommand("SELECT COUNT(distinct(periodo)) FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);
        //    numeroMeses = (Int32)db.ExecuteScalar(cmd);
        //    return numeroMeses;
        //}
        public string[] Obtener_YTD_meses(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
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

            sPeriodosYTDMeses[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodosYTDMeses[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);

            cmd = db.GetSqlStringCommand("SELECT COUNT(distinct(periodo)) FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);
            Numero_Meses_en_YTD = (Int32)db.ExecuteScalar(cmd);

            return sPeriodosYTDMeses;
        }
        public string[] Obtener_YTD_meses_One_Year_Ago(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial_YTD = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + Periodos[0].Year + " AND MES = " + Periodos[0].Month + " ORDER BY IDPERIODO DESC");
            xMesFin_YTD = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_YTD + " AND IDPERIODO <= " + xMesFin_YTD);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
                    //sCabecera48Meses[indice] = reader[0].ToString();
                    indice++;
                }
            }

            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial_YTD + " AND IDPERIODO <= " + xMesFin_YTD);
            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                while (reader.Read())
                {
                    xPeriodosInt.Append(reader[0].ToString() + ",");
                    Debug.WriteLine(reader[0].ToString() + ",");
                }
            }

            sPeriodosYTDMesesOneYearAgo[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            sPeriodosYTDMesesOneYearAgo[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);
            return sPeriodosYTDMesesOneYearAgo;
        }
        public string[] Obtener_Ultimos_48_meses(int pAño, int pMes)
        {
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT IDPERIODO FROM PERIODOS WHERE AÑO = " + pAño + " AND MES = " + pMes + " ORDER BY IDPERIODO ASC");
            xMesInicial = (Int32)(db.ExecuteScalar(cmd));

            cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM PERIODOS WHERE IDPERIODO >= " + xMesInicial + " AND IDPERIODO <= " + xMesFin);

            using (IDataReader reader = db.ExecuteReader(cmd))
            {
                int indice = 0;
                while (reader.Read())
                {
                    xPeriodos.Append("[" + reader[0].ToString() + "],");
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
        public string[] Obtener_Ultimos_48_meses_Factores_Capital(int pAñoFin, int pMesFin, int pAñoIni, int pMesIni)
        {
            string xMesInicial;
            string xMesFinal;
            StringBuilder xPeriodos = new StringBuilder();
            StringBuilder xPeriodosInt = new StringBuilder();
            DbCommand cmd;
            cmd = db.GetSqlStringCommand("SELECT PERIODO FROM FACTORES WHERE AÑO = " + pAñoFin + " AND MES = " + pMesFin + " AND IDCIUDADESTUDIO = 1 ORDER BY IDPOBLACION ASC");
            xMesInicial = db.ExecuteScalar(cmd).ToString();
            cmd = db.GetSqlStringCommand("SELECT PERIODO FROM FACTORES WHERE AÑO = " + pAñoIni + " AND MES = " + pMesIni + " AND IDCIUDADESTUDIO = 1 ORDER BY IDPOBLACION ASC");
            xMesFinal = db.ExecuteScalar(cmd).ToString();

            //// GENERAR CABECERAS TABLA FACTORES
            //cmd = db.GetSqlStringCommand("SELECT DISTINCT PERIODO FROM FACTORES WHERE PERIODO >= '" + xMesInicial + "' AND PERIODO <= '" + xMesFinal + "' AND IDCIUDADESTUDIO = 1");
            //using (IDataReader reader = db.ExecuteReader(cmd))
            //{
            //    while (reader.Read())
            //    {
            //        xPeriodos.Append("[" + reader[0].ToString() + "],");
            //    }
            //}
            //// GENERAR PERIODOS TABLA FACTORES
            //cmd = db.GetSqlStringCommand("SELECT IDPOBLACION FROM FACTORES WHERE PERIODO >= '" + xMesInicial + "' AND PERIODO <= '" + xMesFinal + "'");
            //using (IDataReader reader = db.ExecuteReader(cmd))
            //{
            //    while (reader.Read())
            //    {
            //        xPeriodosInt.Append(reader[0].ToString() + ",");                    
            //    }
            //}
            //sPeriodos48MesesFactor[0] = xPeriodos.ToString().Substring(0, xPeriodos.Length - 1);
            //sPeriodos48MesesFactor[1] = xPeriodosInt.ToString().Substring(0, xPeriodosInt.Length - 1);

            sPeriodos48MesesFactor[0] = xMesInicial;
            sPeriodos48MesesFactor[1] = xMesFinal;
            return sPeriodos48MesesFactor;
        }




        //public void Leer_Ultimos_48_Meses(string Cab, string xPeriodos)
        //{            
        //    /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES (CAPITAL Y PROVINCIA) */

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
        //                    sdata48Meses_x_Region_NSE[rows, i] = double.Parse(reader[i].ToString());
        //                    Debug.WriteLine(sCabecera48Meses[i-1]);

        //                    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
        //                    {
        //                        db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, double.Parse(reader[i].ToString()));
        //                        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i - 1]);
        //                        //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
        //                        //db_Zoho.ExecuteNonQuery(cmd_1);
        //                        db_Zoho.ExecuteNonQuery(cmd_1);
        //                    }
        //                }
        //                rows++;
        //            }
        //            for (int i = 1; i < cols; i++)
        //            {
        //                Debug.WriteLine(sdata48Meses_x_Region_NSE[0, i]);
        //                sdata48Meses_x_Region_NSE[2, i] = sdata48Meses_x_Region_NSE[0, i] + sdata48Meses_x_Region_NSE[1, i];
        //                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
        //                {
        //                    db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, sdata48Meses_x_Region_NSE[0, i] + sdata48Meses_x_Region_NSE[1, i]);
        //                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, sCabecera48Meses[i-1]);
        //                    //Debug.Write(double.Parse(reader[i].ToString()) + "/" + reader[i].ToString());
        //                    //db_Zoho.ExecuteNonQuery(cmd_1);
        //                    db_Zoho.ExecuteNonQuery(cmd_1);
        //                }
        //            }
        //        }

        //    }
        //}

        public void Limpiar_BD_Resultados()
        {
            DbCommand cmdTabla;
            cmdTabla = db_Zoho.GetSqlStringCommand("TRUNCATE TABLE ZOHO.dbo.[Base De Datos Zoho]");
            db_Zoho.ExecuteNonQuery(cmdTabla);
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            //double capital = 0;
            //double provinci = 0;

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
                                Mercado_ = "1. Capital";
                            }
                            else
                            {
                                sdata48Meses_x_Capital_Provincia[1, i - 2] = sdata48Meses_x_Capital_Provincia[1, i - 2] + valor;
                                Mercado_ = "2. Ciudades";
                            }
                            // implementar switch para validar NSE
                            
                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 1:
                                    NSE_ = "Alto";
                                    break;
                                case 2:
                                    NSE_ = "Medio";
                                    break;
                                case 3:
                                    NSE_ = "Medio Bajo";
                                    break;
                                case 4:
                                    NSE_ = "Bajo";
                                    break;
                                case 5:
                                    NSE_ = "Muy Bajo";
                                    break;
                            }
                            
                            if (i <= 10)
                            {
                                Periodo = "0" + ((i - 2) + 1) + ". " + sCabecera48Meses[i - 2]; 
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + sCabecera48Meses[i - 2]; 
                            }
                            Actualizar_BD(NSE_, "Suma", "DOLARES", Mercado_, "0. Cosmeticos", "DOLARES (%)",
                                "MENSUAL", Periodo, valor, int.Parse(sCabecera48Meses[i - 2].Substring(0, 4)));
                            //"MENSUAL", sCabecera48Meses[i - 2], valor, int.Parse(sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    /* INSERTANDO VALORES DEL TOTAL CAPITAL A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        //Debug.WriteLine(sdata48Meses_x_Capital_Provincia[0, i] + " - " + sdata48Meses_x_Capital_Provincia[1, i]);   
                        //OBTIENE TOTAL PAIS Y LO ALMACENA EN ARRAY sdata48Meses_x_Total
                        sdata48Meses_x_Total[0, i] = sdata48Meses_x_Capital_Provincia[0, i] + sdata48Meses_x_Capital_Provincia[1, i];

                        Actualizar_BD("Cosmeticos", "Suma", "DOLARES", "1. Capital", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[0, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        // LOS MISMOS VALORES DEL SP ANTERIOR PERO ACTUALIZANDO V1 Y CIUDAD
                        Actualizar_BD("Lima", "Suma", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[0, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                    }
                    
                    /* INSERTANDO VALORES DEL TOTAL PROVINCIA A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        Actualizar_BD("Cosmeticos", "Suma", "DOLARES", "2. Ciudades", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[1, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        // LOS MISMOS VALORES DEL SP ANTERIOR PERO ACTUALIZANDO V1 Y CIUDAD
                        Actualizar_BD("Ciudades", "Suma", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[1, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                    }

                    for (int i = 0; i < sdata48Meses_x_Total.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                        }
                        //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                        /* INSERTANDO VALORES DEL TOTAL PAIS A BD ZOHO*/
                        Actualizar_BD("Consolidado", "Suma", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", 
                            "MENSUAL", Periodo, sdata48Meses_x_Total[0, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                    }

                    for (int id = 0; id < 5; id++)
                    {
                        switch (id)
                        {
                            case 0:
                                NSE_ = "Alto";
                                break;
                            case 1:
                                NSE_ = "Medio";
                                break;
                            case 2:
                                NSE_ = "Medio Bajo";
                                break;
                            case 3:
                                NSE_ = "Bajo";
                                break;
                            case 4:
                                NSE_ = "Muy Bajo";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_NSE.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "Suma", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", Periodo, 
                                sdata48Meses_x_Region_NSE[id, i] + sdata48Meses_x_Region_NSE[id + 5, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y CATEGORIAS (CAPITAL Y PROVINCIA Y CATEGORIAS - FRAGANCIAS,MAQUILLAJE,CUIDADO PERSONAL) */
            //double capital = 0;
            //double provinci = 0;

            string consulta = @"SELECT REGION, PARAMETROCATEGORIA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, PARAMETROCATEGORIA, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 AND CATEG IN ('FRAGANCIAS','MAQUILLAJE','CUIDADO PERSONAL') ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, PARAMETROCATEGORIA";

            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    int cols = reader.FieldCount;
                    int rows = 0;
                    double valor;

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

                            sdata48Meses_x_Region_Categoria_Mes[rows, i - 2] = valor;

                            // implementar switch para validar NSE

                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 123:
                                    NSE_ = "Fragancias";
                                    Mercado_ = "1. Fragancias";
                                    break;
                                case 124:
                                    NSE_ = "Maquillaje";
                                    Mercado_ = "2. Maquillaje";
                                    break;
                                case 127:
                                    NSE_ = "Cuidado Personal";
                                    Mercado_ = "5. Cuidado Personal";
                                    break;
                            }

                            if (reader[0].ToString() == "CAPITAL")
                            {
                                Ciudad_ = "1. Capital";
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                            }

                            if (i <= 10)
                            {
                                Periodo = "0" + ((i - 2) + 1) + ". " + sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "Suma", "DOLARES", Ciudad_, Mercado_, "DOLARES (%)",
                                "MENSUAL", Periodo, valor, int.Parse(sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    for (int id = 0; id < 3; id++)
                    {
                        switch (id)
                        {
                            case 0:
                                NSE_ = "Fragancias";
                                Mercado_ = "1. Fragancias";
                                break;
                            case 1:
                                NSE_ = "Maquillaje";
                                Mercado_ = "2. Maquillaje";
                                break;
                            case 2:
                                NSE_ = "Cuidado Personal";
                                Mercado_= "5. Cuidado Personal";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_Categoria_Mes.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "Suma", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_Mes[id, i] + sdata48Meses_x_Region_Categoria_Mes[id + 3, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y CANAL DE VENTA (CAPITAL Y PROVINCIA Y CANAL DE VENTA) 
                88 - TRADICIONAL
                87 - DIRECTA
                AND CATEG IN ('FRAGANCIAS','MAQUILLAJE','CUIDADO PERSONAL') ; para filtrar por categorias de ser necesario
             */

            string consulta = @"SELECT REGION, COD_MOD_VENTA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, COD_MOD_VENTA, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2  ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, COD_MOD_VENTA";

            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    int cols = reader.FieldCount;
                    int rows = 0;
                    double valor;

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

                            sdata48Meses_x_Region_Modalidad_Mes[rows, i - 2] = valor;

                            // implementar switch para validar NSE

                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 87:
                                    NSE_ = "VD";
                                    Mercado_ = "1. VD";
                                    break;
                                case 88:
                                    NSE_ = "VR";
                                    Mercado_ = "2. VR";
                                    break;
                            }

                            if (reader[0].ToString() == "CAPITAL")
                            {
                                Ciudad_ = "1. Capital";
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                            }

                            if (i <= 10)
                            {
                                Periodo = "0" + ((i - 2) + 1) + ". " + sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "Suma", "DOLARES", Ciudad_, Mercado_, "DOLARES (%)",
                                "MENSUAL", Periodo, valor, int.Parse(sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    for (int id = 0; id < 2; id++)
                    {
                        switch (id)
                        {
                            case 0:
                                NSE_ = "VD";
                                Mercado_ = "1. VD";
                                break;
                            case 1:
                                NSE_ = "VR";
                                Mercado_ = "2. VR";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_Modalidad_Mes.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES POR MODALIDAD A TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "Suma", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Modalidad_Mes[id, i] + sdata48Meses_x_Region_Modalidad_Mes[id + 2, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_TIPOS(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y TIPOS(CAPITAL Y PROVINCIA Y TIPOS) */

            string consulta = @"SELECT REGION, PARAMETROTIPO, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, PARAMETROTIPO, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 AND TIPO IN ('COLONIA FEMENINAS','COLONIA MASCULINAS','HUMECTANTE/NUTRITIVA CORPORAL','NUTRITIVA REVIT. FACIAL','ROLL-ON','SHAMPOO ADULTOS') ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, PARAMETROTIPO";

            using (DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    int cols = reader.FieldCount;
                    int rows = 0;
                    double valor;

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

                            sdata48Meses_x_Region_Tipos_Mes[rows, i - 2] = valor;

                            // implementar switch para validar TIPOS

                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 158:
                                    NSE_ = "Colonia Femeninas";
                                    Mercado_ = "01. Colonia Femeninas";
                                    break;
                                case 161:
                                    NSE_ = "Colonia Masculinas";
                                    Mercado_ = "02. Colonia Masculinas";
                                    break;
                                case 202:
                                    NSE_ = "Nutritiva Revit. Facial";
                                    Mercado_ = "08. Nutritiva Revit. Facial";
                                    break;
                                case 215:
                                    NSE_ = "Humectante/Nutritiva Corporal";
                                    Mercado_ = "09. Humectante/Nutritiva Corporal";
                                    break;
                                case 226:
                                    NSE_ = "Shampoo Adultos";
                                    Mercado_ = "10. Shampoo Adultos";
                                    break;
                                case 237:
                                    NSE_ = "Roll-On";
                                    Mercado_ = "14. Roll-On";
                                    break;
                            }

                            if (reader[0].ToString() == "CAPITAL")
                            {
                                Ciudad_ = "1. Capital";
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                            }

                            if (i <= 10)
                            {
                                Periodo = "0" + ((i - 2) + 1) + ". " + sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "Suma", "DOLARES", Ciudad_, Mercado_, "DOLARES (%)",
                                "MENSUAL", Periodo, valor, int.Parse(sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    for (int id = 0; id < 6; id++)
                    {
                        switch (id)
                        {
                            case 0:
                                NSE_ = "Colonia Femeninas";
                                Mercado_ = "01. Colonia Femeninas";
                                break;
                            case 1:
                                NSE_ = "Colonia Masculinas";
                                Mercado_ = "02. Colonia Masculinas";
                                break;
                            case 2:
                                NSE_ = "Nutritiva Revit. Facial";
                                Mercado_ = "08. Nutritiva Revit. Facial";
                                break;
                            case 3:
                                NSE_ = "Humectante/Nutritiva Corporal";
                                Mercado_ = "09. Humectante/Nutritiva Corporal";
                                break;
                            case 4:
                                NSE_ = "Shampoo Adultos";
                                Mercado_ = "10. Shampoo Adultos";
                                break;
                            case 5:
                                NSE_ = "Roll-On";
                                Mercado_ = "14. Roll-On";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_Tipos_Mes.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES POR TIPOS TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "Suma", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Tipos_Mes[id, i] + sdata48Meses_x_Region_Tipos_Mes[id + 6, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        private void Actualizar_BD(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, string _Periodo, double _Datos, int _Año)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_INSERTA"))
            {
                db_Zoho.AddInParameter(cmd_1, "_V1", DbType.String, _V1);
                db_Zoho.AddInParameter(cmd_1, "_V2", DbType.String, _V2);
                db_Zoho.AddInParameter(cmd_1, "_VARIABLE", DbType.String, _Variable);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, _Ciudad);
                db_Zoho.AddInParameter(cmd_1, "_MERCADO", DbType.String, _Mercado);
                db_Zoho.AddInParameter(cmd_1, "_UNIDAD", DbType.String, _Unidad);
                db_Zoho.AddInParameter(cmd_1, "_REPORTE", DbType.String, _Reporte);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, _Periodo);
                db_Zoho.AddInParameter(cmd_1, "_DATOS", DbType.Decimal, _Datos);
                db_Zoho.AddInParameter(cmd_1, "_AÑO", DbType.Int16, _Año);

                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }



    }
}
