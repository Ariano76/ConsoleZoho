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
    public class BL_HOGARES
    {
        private double[,] sdata48Meses_x_Region_NSE = new double[10, 48];
        private double[,] sdata48Meses_x_Capital_Provincia = new double[2, 48];
        private double[,] sdata48Meses_x_Total = new double[1, 48];
        public double[,] sdata48Meses_x_Region_Categoria_Mes = new double[6, 48];
        public double[,] sdata48Meses_x_Region_Modalidad_Mes = new double[4, 48];

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();

        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        public string resultadoBD;
        string NSE_, Ciudad_ , Mercado_, Periodo ;

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
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, "0.Cosmeticos", "HOGARES",
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
