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

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();

        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        public string resultadoBD;
        string NSE_, Ciudad_ , Mercado_, prueba ;

        public void Leer_Ultimos_48_Meses_Ciudad_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA DE HOGARES EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            //BD_Zoho xObjeto = BD_Zoho();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_CREA_TABLA_TEMPORAL"))
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
                            
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Mercado_, "0. Cosmeticos", "HOGARES",
                                "MENSUAL", BD_Zoho.sCabecera48Meses[i - 2], valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 2].Substring(0, 4)));
                        }
                        rows++;
                    }

                    /* INSERTANDO VALORES DEL TOTAL CAPITAL A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        //OBTIENE TOTAL PAIS Y LO ALMACENA EN ARRAY sdata48Meses_x_Total
                        sdata48Meses_x_Total[0, i] = sdata48Meses_x_Capital_Provincia[0, i] + sdata48Meses_x_Capital_Provincia[1, i];

                        Actualizar_BD("COSMETICOS", "SUMA", "HOGARES", "1. Capital", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", BD_Zoho.sCabecera48Meses[i], sdata48Meses_x_Capital_Provincia[0, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    }

                    /* INSERTANDO VALORES DEL TOTAL PROVINCIA A BD ZOHO*/
                    for (int i = 0; i < sdata48Meses_x_Capital_Provincia.GetLength(1); i++)
                    {
                        Actualizar_BD("COSMETICOS", "SUMA", "HOGARES", "2. Ciudades", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", BD_Zoho.sCabecera48Meses[i], sdata48Meses_x_Capital_Provincia[1, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
                    }

                    for (int i = 0; i < sdata48Meses_x_Total.GetLength(1); i++)
                    {
                        /* INSERTANDO VALORES DEL TOTAL PAIS A BD ZOHO*/
                        Actualizar_BD("CONSOLIDADO", "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES",
                            "MENSUAL", BD_Zoho.sCabecera48Meses[i], sdata48Meses_x_Total[0, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
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
                            /* INSERTANDO VALORES DE NSE TOTAL PAIS A BD ZOHO*/
                            Actualizar_BD(NSE_, "SUMA", "HOGARES", "0. Consolidado", "0. Cosmeticos", "HOGARES", "MENSUAL", BD_Zoho.sCabecera48Meses[i],
                                sdata48Meses_x_Region_NSE[id, i] + sdata48Meses_x_Region_NSE[id + 5, i], int.Parse(BD_Zoho.sCabecera48Meses[i].Substring(0, 4)));
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
