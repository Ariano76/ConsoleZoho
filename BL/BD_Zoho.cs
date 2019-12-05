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
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_Region_NSE = new double[10, 48];
        public double[,] sdata48Meses_x_Capital_Provincia = new double[2, 48];
        public double[,] sdata48Meses_x_Total = new double[1, 48];

        public double[,] sdata48Meses_x_Region_Categoria_Mes = new double[6, 48];
        public double[,] sdata48Meses_x_Region_Modalidad_Mes = new double[4, 48];
        public double[,] sdata48Meses_x_Region_Tipos_Mes = new double[12, 48];


        private readonly DateTime[] Periodos = new DateTime[7];
        string sSource = "Dashboard ZOHO";        
        string NSE_, Ciudad_, Mercado_, Periodo;
        int xMesFin;
        public string resultadoBD;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file.
            return "Data Source=PEDT0108243; Initial Catalog=BIP; User Id=Procesamiento; Password=P@ssw0rd";
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
                            Actualizar_BD(NSE_, "SUMA", "DOLARES", Mercado_, "0. Cosmeticos", "DOLARES (%)",
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

                        Actualizar_BD("Cosmeticos", "SUMA", "DOLARES", "1. Capital", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[0, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        // LOS MISMOS VALORES DEL SP ANTERIOR PERO ACTUALIZANDO V1 Y CIUDAD
                        Actualizar_BD("Lima", "SUMA", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)",
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
                        Actualizar_BD("Cosmeticos", "SUMA", "DOLARES", "2. Ciudades", "0. Cosmeticos", "DOLARES (%)",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[1, i], int.Parse(sCabecera48Meses[i].Substring(0, 4)));
                        // LOS MISMOS VALORES DEL SP ANTERIOR PERO ACTUALIZANDO V1 Y CIUDAD
                        Actualizar_BD("Ciudades", "SUMA", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)",
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
                        Actualizar_BD("Consolidado", "SUMA", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", 
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
                            Actualizar_BD(NSE_, "SUMA", "DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", Periodo, 
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

                            Actualizar_BD(NSE_, "SUMA", "DOLARES", Ciudad_, Mercado_, "DOLARES (%)",
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
                            Actualizar_BD(NSE_, "SUMA", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
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

                            Actualizar_BD(NSE_, "SUMA", "DOLARES", Ciudad_, Mercado_, "DOLARES (%)",
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
                            Actualizar_BD(NSE_, "SUMA", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
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
                                    Mercado_ = "1. Colonia Femeninas";
                                    break;
                                case 161:
                                    NSE_ = "Colonia Masculinas";
                                    Mercado_ = "2. Colonia Masculinas";
                                    break;
                                case 202:
                                    NSE_ = "Nutritiva Revit. Facial";
                                    Mercado_ = "3. Nutritiva Revit. Facial";
                                    break;
                                case 215:
                                    NSE_ = "Humectante/Nutritiva Corporal";
                                    Mercado_ = "4. Humectante/Nutritiva Corporal";
                                    break;
                                case 226:
                                    NSE_ = "Shampoo Adultos";
                                    Mercado_ = "5. Shampoo Adultos";
                                    break;
                                case 237:
                                    NSE_ = "Roll-On";
                                    Mercado_ = "6. Roll-On";
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

                            Actualizar_BD(NSE_, "SUMA", "DOLARES", Ciudad_, Mercado_.ToLowerInvariant(), "DOLARES (%)",
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
                                Mercado_ = "1. Colonia Femeninas";
                                break;
                            case 1:
                                NSE_ = "Colonia Masculinas";
                                Mercado_ = "2. Colonia Masculinas";
                                break;
                            case 2:
                                NSE_ = "Nutritiva Revit. Facial";
                                Mercado_ = "3. Nutritiva Revit. Facial";
                                break;
                            case 3:
                                NSE_ = "Humectante/Nutritiva Corporal";
                                Mercado_ = "4. Humectante/Nutritiva Corporal";
                                break;
                            case 4:
                                NSE_ = "Shampoo Adultos";
                                Mercado_ = "5. Shampoo Adultos";
                                break;
                            case 5:
                                NSE_ = "Roll-On";
                                Mercado_ = "6. Roll-On";
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
                            Actualizar_BD(NSE_, "SUMA", "DOLARES", "0. Consolidado", Mercado_, "DOLARES (%)", "MENSUAL", Periodo,
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
