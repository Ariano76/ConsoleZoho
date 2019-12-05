using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Diagnostics;

namespace BL
{
    public class BL_PPU_Dolares
    {
        public int[] sPeriodoActual = new int[2];
        private int[] Codigo_NSE = { 1,2,3,4,5 };
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor;
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_valor = new double[1, 48];       // TOTAL POR REGION VALORES
        public double[,] sdata48Meses_x_unid = new double[1, 48];       // TOTAL POR REGION UNIDADES
        public double[,] sdata48Meses_x_Region_valor = new double[2, 48];       // TOTAL POR REGION VALORES
        public double[,] sdata48Meses_x_Region_unid = new double[2, 48];        // TOTAL POR REGION UNIDADES
        public double[,] sdata48Meses_x_NSE_valor = new double[5, 48];          // TOTAL POR NSE VALORES
        public double[,] sdata48Meses_x_NSE_unid = new double[5, 48];           // TOTAL POR NSE UNIDADES
        public double[,] sdata48Meses_x_Region_NSE_valor = new double[10, 48];  // TOTAL POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Region_NSE_unid = new double[10, 48];   // TOTAL POR REGION Y NSE UNIDADES

        public double[,] sdata48Meses_x_Region_Modalidad = new double[8, 48];   // TOTAL POR REGION, MODALIDAD
        public double[,] sdata48Meses_x_Modalidad_valor = new double[2, 48];          // TOTAL POR MODALIDAD
        public double[,] sdata48Meses_x_Modalidad_unid = new double[2, 48];          // TOTAL POR MODALIDAD




        public double[,] sdata48Meses_x_Region_NSE = new double[20, 48];
        public double[,] sdata48Meses_x_Total = new double[1, 48];
        public double[,] sdata48Meses_x_Total_valor = new double[1, 48]; // VALORES
        public double[,] sdata48Meses_x_Total_unid = new double[1, 48];  // UNIDADES
        public double[,] sdata48Meses_x_Region_Categoria_Mes = new double[6, 48];       
        public double[,] sdata48Meses_x_Region_Tipos_Mes = new double[12, 48];

        private readonly DateTime[] Periodos = new DateTime[7];
        string NSE_, V1, Mercado_, Periodo, Ciudad_;
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

        public void Leer_Ultimos_48_Meses_CIUDAD_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1; 

            string consulta_1 = @"SELECT REGION, NSE, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, FACTOR_UNIDAD_VALORIZADO/1000000 AS FACTOR_UNIDAD_VALORIZADO FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD_VALORIZADO) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION,NSE";

            string consulta_2 = @"SELECT REGION, NSE, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, FACTOR_UNIDAD/1000000 AS FACTOR_UNIDAD FROM BASE_REGIONES " +
                "WHERE IDPERIODO IN (" + @xPeriodos + ") AND IDMONEDA = 2 ) AS SourceTable " +
                "PIVOT (SUM(FACTOR_UNIDAD) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION,NSE";

            using (DbCommand cmd_1 = db.GetSqlStringCommand(consulta_1)) // RUTINA PARA LEER VALORES
            {
                using (IDataReader reader_1 = db.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
                        for (int i = 2; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Region_NSE_valor[rows, i - 2] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db.GetSqlStringCommand(consulta_2)) // RUTINA PARA LEER UNIDADES
            {
                using (IDataReader reader_1 = db.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
                        for (int i = 2; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Region_NSE_unid[rows, i - 2] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Region_NSE_valor.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if ( i < 5)
                    {
                        Mercado_ = "1. Capital";
                        switch (i)
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
                        sdata48Meses_x_NSE_valor[i, x] = sdata48Meses_x_Region_NSE_valor[i, x];
                        sdata48Meses_x_NSE_unid[i, x] = sdata48Meses_x_Region_NSE_unid[i, x];

                        sdata48Meses_x_Region_valor[0, x] = sdata48Meses_x_Region_valor[0, x] + sdata48Meses_x_Region_NSE_valor[i, x];
                        sdata48Meses_x_Region_unid[0, x] = sdata48Meses_x_Region_unid[0, x] + sdata48Meses_x_Region_NSE_unid[i, x];
                    }
                    else
                    {
                        Mercado_ = "2. Ciudades";
                        switch (i)
                        {
                            case 5:
                                NSE_ = "Alto";
                                break;
                            case 6:
                                NSE_ = "Medio";
                                break;
                            case 7:
                                NSE_ = "Medio Bajo";
                                break;
                            case 8:
                                NSE_ = "Bajo";
                                break;
                            case 9:
                                NSE_ = "Muy Bajo";
                                break;
                        }
                        sdata48Meses_x_NSE_valor[i - 5, x] = sdata48Meses_x_NSE_valor[i - 5, x] + sdata48Meses_x_Region_NSE_valor[i, x];
                        sdata48Meses_x_NSE_unid[i - 5, x] = sdata48Meses_x_NSE_unid[i - 5, x] + sdata48Meses_x_Region_NSE_unid[i, x];

                        sdata48Meses_x_Region_valor[1, x] = sdata48Meses_x_Region_valor[1, x] + sdata48Meses_x_Region_NSE_valor[i, x];
                        sdata48Meses_x_Region_unid[1, x] = sdata48Meses_x_Region_unid[1, x] + sdata48Meses_x_Region_NSE_unid[i, x];
                    }
                    if (x < 9)
                    {
                        Periodo = "0" + ( x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_Region_NSE_valor[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Region_NSE_valor[i, x] / sdata48Meses_x_Region_NSE_unid[i, x];
                    }

                    Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", Mercado_, "0. Cosmeticos", "PPU (DOL)",
                        "MENSUAL", Periodo, valor_1, 
                        int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }

            /* INSERTANDO VALORES - PPU TOTAL NSE A BD */
            for (int i = 0; i < 5; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_NSE_valor.GetLength(1); x++) // LEYENDO LAS COLUMNAS DEL ARRAY
                {
                    switch (i)
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

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_NSE_valor[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_NSE_valor[i, x] / sdata48Meses_x_NSE_unid[i, x];
                    }

                    Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", "0. Consolidado", "0. Cosmeticos", "PPU (DOL)",
                        "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
            
            /* INSERTANDO VALORES - PPU TOTAL REGION A BD */
            for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Region_valor.GetLength(1); x++) // LEYENDO LAS COLUMNAS DEL ARRAY
                {
                    switch (i)
                    {
                        case 0:
                            Mercado_ = "1. Capital";
                            V1 = "Lima";
                            break;
                        case 1:
                            Mercado_ = "2. Ciudades";
                            V1 = "Ciudades";
                            break;
                    }
                    sdata48Meses_x_valor[0, x] = sdata48Meses_x_valor[0, x] + sdata48Meses_x_Region_valor[i, x];
                    sdata48Meses_x_unid[0, x] = sdata48Meses_x_unid[0, x] + sdata48Meses_x_Region_unid[i, x];

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_Region_valor[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Region_valor[i, x] / sdata48Meses_x_Region_unid[i, x];
                    }

                    Actualizar_BD("Cosmeticos", "SUMA", "PPU (DOL)", Mercado_, "0. Cosmeticos", "PPU (DOL)",
                        "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

                    // COPIAN VALORES CON VARIABLES INVERTIDAS
                    Actualizar_BD(V1, "SUMA", "PPU (DOL)", "0. Consolidado", "0. Cosmeticos", "PPU (DOL)",
                        "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }

            /* INSERTANDO VALORES - PPU TOTAL PAIS A BD */
            for (int x = 0; x < sdata48Meses_x_valor.GetLength(1); x++) // LEYENDO LAS COLUMNAS DEL ARRAY
            {
                if (x < 9)
                {
                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                }
                else
                {
                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                }
                if (sdata48Meses_x_valor[0, x] <= 0)
                {
                    valor_1 = 0;
                }
                else
                {
                    valor_1 = sdata48Meses_x_valor[0, x] / sdata48Meses_x_unid[0, x];
                }
                Actualizar_BD("Consolidado", "SUMA", "PPU (DOL)", "0. Consolidado", "0. Cosmeticos", "PPU (DOL)",
                        "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
            }

        }

        public void Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_TABLA_TEMPORAL_PPU_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
                        for (int i = 3; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Region_Modalidad[rows, i - 3] = valor_1;
                        }
                        rows++;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                for (int i = 0; i < 4; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 | i == 1 | i == 4 | i == 5)
                        {
                            Ciudad_ = "1. Capital";
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                        }
                        if (i == 0 | i == 2 ) // SE SUMAN LOS CANTIDADES POR CANAL SEGUN LA MONEDA FILA 0 Y 2 SON VD.
                        {
                            sdata48Meses_x_Modalidad_valor[0, x] = sdata48Meses_x_Modalidad_valor[0, x] + sdata48Meses_x_Region_Modalidad[i, x];
                            sdata48Meses_x_Modalidad_unid[0, x] = sdata48Meses_x_Modalidad_unid[0, x] + sdata48Meses_x_Region_Modalidad[i + 4, x];
                        }
                        else
                        {
                            sdata48Meses_x_Modalidad_valor[1, x] = sdata48Meses_x_Modalidad_valor[1, x] + sdata48Meses_x_Region_Modalidad[i, x];
                            sdata48Meses_x_Modalidad_unid[1, x] = sdata48Meses_x_Modalidad_unid[1, x] + sdata48Meses_x_Region_Modalidad[i + 4, x];
                        }
                        if (i % 2 == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE_ = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE_ = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Region_NSE_valor[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Region_Modalidad[i, x] / sdata48Meses_x_Region_Modalidad[i + 4, x];
                        }

                        Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", Ciudad_, "0. Cosmeticos", "PPU (DOL)",
                            "MENSUAL", Periodo, valor_1,
                            int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

                        Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", Ciudad_, Mercado_, "PPU (DOL)",
                            "MENSUAL", Periodo, valor_1,
                            int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
                // CALCULANDO PPU TOTAL MERCADO

                //for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                //{
                //    for (int x = 0; x < sdata48Meses_x_Modalidad_valor.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                //    {
                //        Debug.Print(sdata48Meses_x_Modalidad_valor[i,x].ToString());
                //    }
                //}

                for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Modalidad_valor.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i % 2 == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE_ = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE_ = "VR";
                        }
                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Modalidad_valor[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Modalidad_valor[i, x] / sdata48Meses_x_Modalidad_unid[i, x];
                        }

                        Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", "0. Consolidado", Mercado_, "PPU (DOL)",
                            "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

                        Actualizar_BD(NSE_, "SUMA", "PPU (DOL)", "0. Consolidado", "0. Cosmeticos", "PPU (DOL)",
                            "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

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
