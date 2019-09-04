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

        private readonly DateTime[] Periodos = new DateTime[7];
        string sSource = "Dashboard ZOHO";
        string sLog = "ZOHO";
        string sEvent = "Mensaje de nuestro evento";
        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");

        static void Main()
        {
            
            
        }

        static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file.
            return "Data Source=PEDT0108243; Initial Catalog=BIP; Integrated Security=SSPI";
        }

        public void Periodo_Actual(int pAño, string pMes)
        {
            //'RECUPERA EL ID DEL PERIODO CORRESPONDIENTE AL AÑO Y MES ANTERIOR
            

            if (!EventLog.SourceExists(sSource))
            {
                EventLog.CreateEventSource(sSource, sLog);
            }

            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                string Sql;
                Sql = "EXEC _SP_ID_PERIODO_AÑO " + pAño + "," + pMes;

                //SqlDataAdapter adapter = new SqlDataAdapter();
                //adapter.TableMappings.Add("Table", "Periodos");

                SqlCommand cmd = new SqlCommand("_SP_ID_PERIODO_AÑO", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@Año", SqlDbType.Int);
                cmd.Parameters["@Año"].Value = pAño;

                cmd.Parameters.Add("@Mes", SqlDbType.VarChar);
                cmd.Parameters["@Mes"].Value = pMes;
                try
                {
                    con.Open();
                    sPeriodoActual[0] = (Int32)cmd.ExecuteScalar();
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

        public string IdMarca()
        {
            string x="";
            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {

            }
            return x;
        }
    }
}
