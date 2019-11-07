using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Configuration;
using Dapper;

namespace BL
{
    public class BL_HOGARES
    {
        // SE TIENE QUE ACTUALIZAR LOS CODIGOS DE LOS TIPOS A ANALIZAR DE ESTE ARREGLO POR LOS DEFINITIVOS
        private int[] codigo_Tipos_Importantes = {158, 161, 215, 202, 237, 226 };

        private double[,] sdata48Meses_x_Region_NSE = new double[10, 48];
        private double[,] sdata48Meses_x_Capital_Provincia = new double[2, 48];
        private double[,] sdata48Meses_x_Total = new double[1, 48];
        public double[,] sdata48Meses_x_Region_Categoria_Mes = new double[6, 48];
        public double[,] sdata48Meses_x_Region_Modalidad_Mes = new double[4, 48];

        public double[,] sdata48Meses_x_Region_Categoria_NSE_Mes = new double[30, 50];
        public double[,] sdata48Meses_x_Region_Categoria_Modalidad_Mes = new double[12, 50];
        public double[,] sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes = new double[60, 50];

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");
        Database db_Zoho_1 = factory.Create("ZOHO");

        public string resultadoBD;
        string V1_, NSE_, Ciudad_ , Mercado_, MercadoV1_, Periodo ;

        public void Leer_Ultimos_48_Meses_Ciudad_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA DE HOGARES EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            //BD_Zoho xObjeto = BD_Zoho();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, NSE, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP15) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
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
                                    NSE_ = "ALTO";
                                    break;
                                case 2:
                                    NSE_ = "MEDIO";
                                    break;
                                case 3:
                                    NSE_ = "MEDIO BAJO";
                                    break;
                                case 4:
                                    NSE_ = "BAJO";
                                    break;
                                case 5:
                                    NSE_ = "MUY BAJO";
                                    break;
                            }

                            if (i <= 10)
                            {
                                Periodo = "0" + ((i - 2) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Mercado_, "0. Cosmeticos", "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    /* INSERTANDO VALORES DEL TOTAL CAPITAL A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        //OBTIENE TOTAL PAIS Y LO ALMACENA EN ARRAY sdata48Meses_x_Total
                        sdata48Meses_x_Total[0, i] = sdata48Meses_x_Capital_Provincia[0, i] + sdata48Meses_x_Capital_Provincia[1, i];

                        Actualizar_BD("COSMETICOS", "SUMA", "HOGARES", "1. Capital", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[0, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                        
                        // GUARDANDO VALORES CON VARIABLES INVERTIDAS - CAPITAL
                        Actualizar_BD("Lima", "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[0, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    }

                    /* INSERTANDO VALORES DEL TOTAL PROVINCIA A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        Actualizar_BD("COSMETICOS", "SUMA", "HOGARES", "2. Ciudades", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[1, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                        
                        // GUARDANDO VALORES CON VARIABLES INVERTIDAS - CIUDADES
                        Actualizar_BD("Ciudades", "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", Periodo, sdata48Meses_x_Capital_Provincia[1, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    }

                    for (int i = 0; i < sdata48Meses_x_Total.GetLength(1); i++)
                    {
                        if (i < 9)
                        {
                            Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        else
                        {
                            Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                        }
                        /* INSERTANDO VALORES DEL TOTAL PAIS A BD ZOHO*/
                        Actualizar_BD("CONSOLIDADO", "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", Periodo, sdata48Meses_x_Total[0, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    }

                    for (int id = 0; id < 5; id++)
                    {
                        switch (id)
                        {
                            case 0:
                                NSE_ = "ALTO";
                                break;
                            case 1:
                                NSE_ = "MEDIO";
                                break;
                            case 2:
                                NSE_ = "MEDIO BAJO";
                                break;
                            case 3:
                                NSE_ = "BAJO";
                                break;
                            case 4:
                                NSE_ = "MUY BAJO";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_NSE.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_NSE[id, i] + sdata48Meses_x_Region_NSE[id + 5, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
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
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_CATEGORIA"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, PARAMETROCATEGORIA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, PARAMETROCATEGORIA, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_CATEGORIA) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
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
                                Periodo = "0" + ((i - 2) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
                            // CREANDO UNA COPIA DE LOS DATOS PERO CON VALOR "0. Consolidado"  EN LA VARIABLE MERCADO
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, "0. Cosmeticos", "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
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
                                Mercado_ = "5. Cuidado Personal";
                                break;
                        }

                        for (int i = 0; i < sdata48Meses_x_Region_Categoria_Mes.GetLength(1); i++)
                        {
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_Mes[id, i] + sdata48Meses_x_Region_Categoria_Mes[id + 3, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                            // CREANDO UNA COPIA DE LOS DATOS PERO CON VALOR "0. Consolidado"  EN LA VARIABLE MERCADO
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_Mes[id, i] + sdata48Meses_x_Region_Categoria_Mes[id + 3, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
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
             */

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, COD_MOD_VENTA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, COD_MOD_VENTA, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_MODALIDAD) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
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

                            // implementar switch para validar CANAL

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
                                Periodo = "0" + ((i - 2) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }
                            else
                            {
                                Periodo = (i - 1) + ". " + BD_Zoho.sCabecera48Meses[i - 2];
                            }

                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
                            // INSERTANDO DATOS DONDE LOS VALORES EN MODALIDAD APARECE COMO '0. Cosmeticos'
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, "0. Cosmeticos", "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
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
                                Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES POR MODALIDAD A TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Modalidad_Mes[id, i] + sdata48Meses_x_Region_Modalidad_Mes[id + 2, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                            // INSERTANDO DATOS DONDE LOS VALORES EN MODALIDAD APARECE COMO '0. Cosmeticos'
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Modalidad_Mes[id, i] + sdata48Meses_x_Region_Modalidad_Mes[id + 2, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y CATEGORIAS (CAPITAL Y PROVINCIA Y CATEGORIAS - FRAGANCIAS,MAQUILLAJE,CUIDADO PERSONAL) */
            //double capital = 0;
            //double provinci = 0;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_CATEGORIA_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, NSE, PARAMETROCATEGORIA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, PARAMETROCATEGORIA, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_CATEGORIA_NSE) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, NSE, PARAMETROCATEGORIA";

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
                        for (int i = 3; i < cols; i++)
                        {
                            if (reader[i] == DBNull.Value)
                            {
                                valor = 0;
                            }
                            else
                            {
                                valor = double.Parse(reader[i].ToString());
                            }
                            
                            sdata48Meses_x_Region_Categoria_NSE_Mes[rows, 0] = int.Parse(reader[1].ToString()); // GUARDANDO NSE
                            sdata48Meses_x_Region_Categoria_NSE_Mes[rows, 1] = int.Parse(reader[2].ToString()); // GUARDANDO PARAMETROCATEGORIA
                            sdata48Meses_x_Region_Categoria_NSE_Mes[rows, i - 1] = valor; // GUARDANDO NUMERO DE HOGARES

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

                            switch (int.Parse(reader[2].ToString()))
                            {
                                case 123:
                                    Mercado_ = "1. Fragancias";
                                    break;
                                case 124:
                                    Mercado_ = "2. Maquillaje";
                                    break;
                                case 127:
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

                            if (i <= 11)
                            {
                                Periodo = "0" + ((i - 3) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }
                            else
                            {
                                Periodo = (i - 2) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }

                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                        }
                        rows++;
                    }

                    for (int id = 0; id < 15; id++)
                    {
                        for (int i = 0; i < sdata48Meses_x_Region_Categoria_NSE_Mes.GetLength(1) - 2; i++)
                        {
                            switch (int.Parse(sdata48Meses_x_Region_Categoria_NSE_Mes[id, 0].ToString()))
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

                            switch (int.Parse(sdata48Meses_x_Region_Categoria_NSE_Mes[id, 1].ToString()))
                            {
                                case 123:
                                    Mercado_ = "1. Fragancias";
                                    break;
                                case 124:
                                    Mercado_ = "2. Maquillaje";
                                    break;
                                case 127:
                                    Mercado_ = "5. Cuidado Personal";
                                    break;
                            }
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_NSE_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_NSE_Mes[id + 15, i + 2],
                                int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }


                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y CATEGORIAS (CAPITAL Y PROVINCIA Y CATEGORIAS - FRAGANCIAS,MAQUILLAJE,CUIDADO PERSONAL) */
            //double capital = 0;
            //double provinci = 0;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_CATEGORIA_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, COD_MOD_VENTA, PARAMETROCATEGORIA, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, COD_MOD_VENTA, PARAMETROCATEGORIA, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_CATEGORIA_MODALIDAD) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, COD_MOD_VENTA, PARAMETROCATEGORIA";

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
                        for (int i = 3; i < cols; i++)
                        {
                            if (reader[i] == DBNull.Value)
                            {
                                valor = 0;
                            }
                            else
                            {
                                valor = double.Parse(reader[i].ToString());
                            }

                            sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, 0] = int.Parse(reader[1].ToString()); // GUARDANDO MODALIDAD
                            sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, 1] = int.Parse(reader[2].ToString()); // GUARDANDO PARAMETROCATEGORIA
                            sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, i - 1] = valor; // GUARDANDO NUMERO DE HOGARES

                            // implementar switch para validar NSE

                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 87:
                                    NSE_ = "VD";
                                    MercadoV1_ = "1. VD";
                                    break;
                                case 88:
                                    NSE_ = "VR";
                                    MercadoV1_ = "2. VR";
                                    break;
                            }

                            switch (int.Parse(reader[2].ToString()))
                            {
                                case 123:
                                    Mercado_ = "1. Fragancias";
                                    V1_ = "Fragancias";
                                    break;
                                case 124:
                                    Mercado_ = "2. Maquillaje";
                                    V1_ = "Maquillaje";
                                    break;
                                case 127:
                                    Mercado_ = "5. Cuidado Personal";
                                    V1_ = "Cuidado Personal";
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

                            if (i <= 11)
                            {
                                Periodo = "0" + ((i - 3) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }
                            else
                            {
                                Periodo = (i - 2) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }

                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                            // DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                            Actualizar_BD(V1_, "SUMA", "HOGARES", Ciudad_, MercadoV1_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                        }
                        rows++;
                    }

                    for (int id = 0; id < 6; id++)
                    {
                        for (int i = 0; i < sdata48Meses_x_Region_Categoria_Modalidad_Mes.GetLength(1) - 2; i++)
                        {
                            switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 0].ToString()))
                            {
                                case 87:
                                    NSE_ = "VD";
                                    MercadoV1_ = "1. VD";
                                    break;
                                case 88:
                                    NSE_ = "VR";
                                    MercadoV1_ = "2. VR";
                                    break;
                            }

                            switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 1].ToString()))
                            {
                                case 123:
                                    Mercado_ = "1. Fragancias";
                                    V1_ = "Fragancias";
                                    break;
                                case 124:
                                    Mercado_ = "2. Maquillaje";
                                    V1_ = "Maquillaje";
                                    break;
                                case 127:
                                    Mercado_ = "5. Cuidado Personal";
                                    V1_ = "Cuidado Personal";
                                    break;
                            }
                            if (i < 9)
                            {
                                Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            else
                            {
                                Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                            }
                            //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                                int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                            // DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                            Actualizar_BD(V1_, "SUMA", "HOGARES", "0. Consolidado", MercadoV1_, "HOGARES", "MENSUAL", Periodo,
                                sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                                int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                        }
                    }


                }
            }
        }




        public void Leer_Ultimos_48_Meses_CIUDAD_NSE_TIPOS(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES, NSE Y TIPOS (CAPITAL Y PROVINCIA, NSE Y TIPOS PRINCIPALES) */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_NSE_TIPOS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, NSE, PARAMETROTIPO, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, PARAMETROTIPO, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_NSE_TIPOS) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, NSE, PARAMETROTIPO";

            int[] codigo_Tipos_Importantes_Completos = new int[6];
            int[] codigo_Tipos_Importantes_Incompletos = new int[6];

            using(DbCommand cmd = db.GetSqlStringCommand(consulta))
            {
                using (IDataReader reader = db.ExecuteReader(cmd))
                {
                    DataTable dtTipos = new DataTable("Tipos");
                    dtTipos.Load(reader);                   
                    int contadorX = 0, contadorY = 0;
                    foreach (int item in codigo_Tipos_Importantes)
                    {
                        DataRow[] tiposFiltrados = dtTipos.Select("PARAMETROTIPO=" + item.ToString());
                        if (Contar_Registros_DT(tiposFiltrados) == true)
                        {
                            codigo_Tipos_Importantes_Completos[contadorX] = item;
                            contadorX++;
                        }
                        else
                        {
                            codigo_Tipos_Importantes_Incompletos[contadorY] = item;
                            contadorY++;
                        }
                    }
                }

                string xTiposCompletos = "";

                //foreach (int item in codigo_Tipos_Importantes_Completos)
                //{
                //    xTiposCompletos += item + ",";
                //}
                for (int i = 0; i < codigo_Tipos_Importantes_Completos.GetLength(0); i++)
                {
                    if (codigo_Tipos_Importantes_Completos[i] != 0)
                    {
                        xTiposCompletos += codigo_Tipos_Importantes_Completos[i] + ",";
                    }
                }

                xTiposCompletos = xTiposCompletos.Substring(0, xTiposCompletos.Length - 1);
                // RUTINA PARA RECORRER LOS TIPOS QUE TIENEN LOS NSE COMPLETOS, 10 REGISTROS.
                consulta = @"SELECT REGION, NSE, PARAMETROTIPO, " + @Cab +
                    " FROM ( SELECT PERIODO, REGION, NSE, PARAMETROTIPO, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_NSE_TIPOS) AS SourceTable " +
                    "PIVOT (SUM(FACTOR_HOGAR) " +
                    "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                    "WHERE PARAMETROTIPO IN (" + xTiposCompletos + ") " +
                    "ORDER BY REGION, NSE, PARAMETROTIPO";

                using (DbCommand cmd_2 = db.GetSqlStringCommand(consulta))
                {
                    using (IDataReader reader_1 = db.ExecuteReader(cmd_2))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        double valor;

                        while (reader_1.Read())
                        {
                            //Debug.Write(reader[0].ToString());
                            /* INICIA EN 2 PARA NO LEER COLUMNA REGION Y NSE*/
                            for (int i = 3; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor = 0;
                                }
                                else
                                {
                                    valor = double.Parse(reader_1[i].ToString());
                                }

                                sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, 0] = int.Parse(reader_1[1].ToString()); // GUARDANDO MODALIDAD
                                sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, 1] = int.Parse(reader_1[2].ToString()); // GUARDANDO PARAMETROCATEGORIA
                                sdata48Meses_x_Region_Categoria_Modalidad_Mes[rows, i - 1] = valor; // GUARDANDO NUMERO DE HOGARES

                                // implementar switch para validar NSE
                                switch (int.Parse(reader_1[1].ToString()))
                                {
                                    case 1:
                                        NSE_ = "Alto";
                                        MercadoV1_ = "1. VD";
                                        break;
                                    case 2:
                                        NSE_ = "Medio";
                                        MercadoV1_ = "2. VR";
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

                                switch (int.Parse(reader_1[2].ToString()))
                                {
                                    case 158:
                                        Mercado_ = "01. Colonia Femeninas";
                                        V1_ = "Fragancias";
                                        break;
                                    case 161:
                                        Mercado_ = "02. Colonia Masculinas";
                                        V1_ = "Maquillaje";
                                        break;
                                    case 215:
                                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 202:
                                        Mercado_ = "08. Nutritiva Revit. Facial";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 237:
                                        Mercado_ = "14. Roll-On";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 226:
                                        Mercado_ = "10. Shampoo Adultos";
                                        V1_ = "Cuidado Personal";
                                        break;
                                }

                                if (reader_1[0].ToString() == "CAPITAL")
                                {
                                    Ciudad_ = "1. Capital";
                                }
                                else
                                {
                                    Ciudad_ = "2. Ciudades";
                                }

                                if (i <= 11)
                                {
                                    Periodo = "0" + ((i - 3) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                                }
                                else
                                {
                                    Periodo = (i - 2) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                                }

                                Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                    "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                                ////DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                                ////Actualizar_BD(V1_, "SUMA", "HOGARES", Ciudad_, MercadoV1_, "HOGARES",
                                ////    "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                            }
                            rows++;
                        }

                        for (int id = 0; id < 6; id++)
                        {
                            for (int i = 0; i < sdata48Meses_x_Region_Categoria_Modalidad_Mes.GetLength(1) - 2; i++)
                            {
                                switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 0].ToString()))
                                {
                                    case 1:
                                        NSE_ = "Alto";
                                        MercadoV1_ = "1. VD";
                                        break;
                                    case 2:
                                        NSE_ = "Medio";
                                        MercadoV1_ = "2. VR";
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

                                switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 1].ToString()))
                                {
                                    case 158:
                                        Mercado_ = "01. Colonia Femeninas";
                                        V1_ = "Fragancias";
                                        break;
                                    case 161:
                                        Mercado_ = "02. Colonia Masculinas";
                                        V1_ = "Maquillaje";
                                        break;
                                    case 215:
                                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 202:
                                        Mercado_ = "08. Nutritiva Revit. Facial";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 237:
                                        Mercado_ = "14. Roll-On";
                                        V1_ = "Cuidado Personal";
                                        break;
                                    case 226:
                                        Mercado_ = "10. Shampoo Adultos";
                                        V1_ = "Cuidado Personal";
                                        break;
                                }
                                if (i < 9)
                                {
                                    Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                                }
                                else
                                {
                                    Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                                }
                                //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                                /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                                Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                                    sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                                    int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                                //////DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                                //////Actualizar_BD(V1_, "SUMA", "HOGARES", "0. Consolidado", MercadoV1_, "HOGARES", "MENSUAL", Periodo,
                                //////    sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                                //////    int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                            }
                        }
                    }
                }
            }
        }

        private bool Contar_Registros_DT(DataRow[] dataRows)
        {
            bool respuesta = true;

            if (dataRows.Count()<10)
            {
                respuesta = false;
            }
            return respuesta;
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_NSE_TIPOS_V1(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y CATEGORIAS (CAPITAL Y PROVINCIA Y CATEGORIAS - FRAGANCIAS,MAQUILLAJE,CUIDADO PERSONAL) */
            //double capital = 0;
            //double provinci = 0;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL_HOGAR_NSE_TIPOS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

            string consulta = @"SELECT REGION, NSE, PARAMETROTIPO, " + @Cab +
                " FROM ( SELECT PERIODO, REGION, NSE, PARAMETROTIPO, FACTOR_HOGAR/1000 AS FACTOR_HOGAR FROM ##TEMP_HOGAR_NSE_TIPOS) AS SourceTable " +
                "PIVOT (SUM(FACTOR_HOGAR) " +
                "FOR PERIODO IN (" + @Cab + ")) AS pvt " +
                "ORDER BY REGION, NSE, PARAMETROTIPO";

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
                        for (int i = 3; i < cols; i++)
                        {
                            if (reader[i] == DBNull.Value)
                            {
                                valor = 0;
                            }
                            else
                            {
                                valor = double.Parse(reader[i].ToString());
                            }

                            sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes[rows, 0] = int.Parse(reader[1].ToString()); // GUARDANDO MODALIDAD
                            sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes[rows, 1] = int.Parse(reader[2].ToString()); // GUARDANDO PARAMETROCATEGORIA
                            sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes[rows, i - 1] = valor; // GUARDANDO NUMERO DE HOGARES

                            // implementar switch para validar NSE

                            switch (int.Parse(reader[1].ToString()))
                            {
                                case 1:
                                    NSE_ = "Alto";
                                    MercadoV1_ = "1. VD";
                                    break;
                                case 2:
                                    NSE_ = "Medio";
                                    MercadoV1_ = "2. VR";
                                    break;
                                case 3:
                                    NSE_ = "Medio Bajo";
                                    MercadoV1_ = "2. VR";
                                    break;
                                case 4:
                                    NSE_ = "Bajo";
                                    MercadoV1_ = "2. VR";
                                    break;
                                case 5:
                                    NSE_ = "Muy Bajo";
                                    MercadoV1_ = "2. VR";
                                    break;
                            }

                            switch (int.Parse(reader[2].ToString()))
                            {
                                case 158:
                                    Mercado_ = "Colonia Femeninas";
                                    V1_ = "1. COLONIA FEMENINAS";
                                    break;
                                case 161:
                                    Mercado_ = "Colonia Masculinas";
                                    V1_ = "2. COLONIA MASCULINAS";
                                    break;
                                case 202:
                                    Mercado_ = "Nutritiva Revit. Facial";
                                    V1_ = "3. NUTRITIVA REVIT. FACIAL";
                                    break;
                                case 215:
                                    Mercado_ = "Humectante/Nutritiva Corporal";
                                    V1_ = "4. HUMECTANTE/NUTRITIVA CORPORAL";
                                    break;
                                case 226:
                                    Mercado_ = "Shampoo Adultos";
                                    V1_ = "5. SHAMPOO ADULTOS";
                                    break;
                                case 237:
                                    Mercado_ = "Roll-On";
                                    V1_ = "6. ROLL-ON";
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

                            if (i <= 11)
                            {
                                Periodo = "0" + ((i - 3) + 1) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }
                            else
                            {
                                Periodo = (i - 2) + ". " + BD_Zoho.sCabecera48Meses[i - 3];
                            }

                            Add_Dapper(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES",
                                "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));

                            // DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                            //Actualizar_BD(V1_, "SUMA", "HOGARES", Ciudad_, MercadoV1_, "HOGARES",
                            //    "MENSUAL", Periodo, valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
                        }
                        rows++;
                    }

                    //for (int id = 0; id < 6; id++)
                    //{
                    //    for (int i = 0; i < sdata48Meses_x_Region_Categoria_Modalidad_Mes.GetLength(1) - 2; i++)
                    //    {
                    //        switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 0].ToString()))
                    //        {
                    //            case 87:
                    //                NSE_ = "VD";
                    //                MercadoV1_ = "1. VD";
                    //                break;
                    //            case 88:
                    //                NSE_ = "VR";
                    //                MercadoV1_ = "2. VR";
                    //                break;
                    //        }

                    //        switch (int.Parse(sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, 1].ToString()))
                    //        {
                    //            case 123:
                    //                Mercado_ = "1. Fragancias";
                    //                V1_ = "Fragancias";
                    //                break;
                    //            case 124:
                    //                Mercado_ = "2. Maquillaje";
                    //                V1_ = "Maquillaje";
                    //                break;
                    //            case 127:
                    //                Mercado_ = "5. Cuidado Personal";
                    //                V1_ = "Cuidado Personal";
                    //                break;
                    //        }
                    //        if (i < 9)
                    //        {
                    //            Periodo = "0" + (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                    //        }
                    //        else
                    //        {
                    //            Periodo = (i + 1) + ". " + BD_Zoho.sCabecera48Meses[i];
                    //        }
                    //        //Debug.WriteLine(sdata48Meses_x_Total[0, i]);                       
                    //        /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                    //        Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", Mercado_, "HOGARES", "MENSUAL", Periodo,
                    //            sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                    //            int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    //        // DUPLICANDO VALORES CON VARIABLES INVERTIDAS
                    //        Actualizar_BD(V1_, "SUMA", "HOGARES", "0. Consolidado", MercadoV1_, "HOGARES", "MENSUAL", Periodo,
                    //            sdata48Meses_x_Region_Categoria_Modalidad_Mes[id, i + 2] + sdata48Meses_x_Region_Categoria_Modalidad_Mes[id + 6, i + 2],
                    //            int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    //    }
                    //}
                }
            }
        }


        private void Actualizar_BD(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, string _Periodo, double _Datos, int _Año)
        {
            using (DbCommand cmd_2 = db_Zoho_1.GetStoredProcCommand("_SP_INSERTA"))
            {                
                db_Zoho_1.AddInParameter(cmd_2, "_V1", DbType.String, _V1);
                db_Zoho_1.AddInParameter(cmd_2, "_V2", DbType.String, _V2);
                db_Zoho_1.AddInParameter(cmd_2, "_VARIABLE", DbType.String, _Variable);
                db_Zoho_1.AddInParameter(cmd_2, "_CIUDAD", DbType.String, _Ciudad);
                db_Zoho_1.AddInParameter(cmd_2, "_MERCADO", DbType.String, _Mercado);
                db_Zoho_1.AddInParameter(cmd_2, "_UNIDAD", DbType.String, _Unidad);
                db_Zoho_1.AddInParameter(cmd_2, "_REPORTE", DbType.String, _Reporte);
                db_Zoho_1.AddInParameter(cmd_2, "_PERIODO", DbType.String, _Periodo);
                db_Zoho_1.AddInParameter(cmd_2, "_DATOS", DbType.Decimal, _Datos);
                db_Zoho_1.AddInParameter(cmd_2, "_AÑO", DbType.Int32, _Año);
                db_Zoho_1.ExecuteNonQuery(cmd_2);                
            }
        }

        private void Insert_BD(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, string _Periodo, double _Datos, int _Año)
        {
            SqlConnection PubsConn = new SqlConnection(ConfigurationManager.ConnectionStrings["Zoho"].ConnectionString);
            using (SqlCommand testCMD = new SqlCommand("_SP_INSERTA", PubsConn))
            {
                testCMD.CommandType = CommandType.StoredProcedure;

                SqlParameter V1 = testCMD.Parameters.Add("_V1", SqlDbType.NVarChar,100);
                V1.Direction = ParameterDirection.Input;
                SqlParameter V2 = testCMD.Parameters.Add("_V2", SqlDbType.NVarChar, 100);
                V2.Direction = ParameterDirection.Input;
                SqlParameter Variable = testCMD.Parameters.Add("_VARIABLE", SqlDbType.NVarChar, 100);
                Variable.Direction = ParameterDirection.Input;
                SqlParameter Ciudad = testCMD.Parameters.Add("_CIUDAD", SqlDbType.NVarChar, 100);
                SqlParameter Mercado = testCMD.Parameters.Add("_MERCADO", SqlDbType.NVarChar, 100);
                SqlParameter Unidad = testCMD.Parameters.Add("_UNIDAD", SqlDbType.NVarChar, 100);
                SqlParameter Reporte = testCMD.Parameters.Add("_REPORTE", SqlDbType.NVarChar, 100);
                SqlParameter Periodo = testCMD.Parameters.Add("_PERIODO", SqlDbType.NVarChar, 100);
                SqlParameter Datos = testCMD.Parameters.Add("_DATOS", SqlDbType.Float, 100);
                SqlParameter Año = testCMD.Parameters.Add("_AÑO", SqlDbType.Int, 100);

                V1.Value = _V1;
                V2.Value = _V2;
                Variable.Value = _Variable;
                Ciudad.Value = _Ciudad;
                Mercado.Value = _Mercado;
                Unidad.Value = _Unidad;
                Reporte.Value = _Reporte;
                Periodo.Value = _Periodo;
                Datos.Value = _Datos;
                Año.Value = _Año;
                try
                {
                    PubsConn.Open();
                    testCMD.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                    throw ex;
                }


            }
        }

        private void Add_Dapper(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, string _Periodo, double _Datos, int _Año)
        {
            using (IDbConnection dbDapper = new SqlConnection(ConfigurationManager.ConnectionStrings["Zoho"].ConnectionString))
            {
                var parameters = new DynamicParameters();
                parameters.Add("@_V1", value: _V1, dbType: DbType.String);
                parameters.Add("@_V2", value: _V2, dbType: DbType.String);
                parameters.Add("@_VARIABLE", value: _Variable, dbType: DbType.String);
                parameters.Add("@_CIUDAD", value: _Ciudad, dbType: DbType.String);
                parameters.Add("@_MERCADO", value: _Mercado, dbType: DbType.String);
                parameters.Add("@_UNIDAD", value: _Unidad, dbType: DbType.String);
                parameters.Add("@_REPORTE", value: _Reporte, dbType: DbType.String);
                parameters.Add("@_PERIODO", value: _Periodo, dbType: DbType.String);
                parameters.Add("@_DATOS", value: _Datos, dbType: DbType.Decimal);
                parameters.Add("@_AÑO", value: _Año, dbType: DbType.Int16);

                dbDapper.Execute("_SP_INSERTA", parameters, commandType: CommandType.StoredProcedure);
            }
        }

    }
}
