using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using Dapper;

namespace BL
{
    public class BL_Hogares_V1
    {
        public double[,] sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes = new double[12, 50];
        string V1_, NSE_, Ciudad_, Mercado_, MercadoV1_, Periodo;
        public double[,] sdata48Meses_x_Region_Categoria_Modalidad_Mes = new double[12, 50];

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");        
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        private IDbConnection db_dapper = new SqlConnection(ConfigurationManager.ConnectionStrings["Zoho"].ConnectionString);

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
                            //sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes[rows, 1] = int.Parse(reader[2].ToString()); // GUARDANDO PARAMETROCATEGORIA
                            //sdata48Meses_x_Region_Categoria_Nse_Tipos_Mes[rows, i - 1] = valor; // GUARDANDO NUMERO DE HOGARES

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
                                    Mercado_ = "COLONIA FEMENINAS";
                                    V1_ = "01. COLONIA FEMENINAS";
                                    break;
                                case 161:
                                    Mercado_ = "COLONIA MASCULINAS";
                                    V1_ = "02. COLONIA MASCULINAS";
                                    break;
                                case 202:
                                    Mercado_ = "NUTRITIVA REVIT. FACIAL";
                                    V1_ = "08. NUTRITIVA REVIT. FACIAL";
                                    break;
                                case 215:
                                    Mercado_ = "HUMECTANTE/NUTRITIVA CORPORAL";
                                    V1_ = "09. HUMECTANTE/NUTRITIVA CORPORAL";
                                    break;
                                case 226:
                                    Mercado_ = "SHAMPOO ADULTOS";
                                    V1_ = "10. SHAMPOO ADULTOS";
                                    break;
                                case 237:
                                    Mercado_ = "ROLL-ON";
                                    V1_ = "14. ROLL-ON";
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


                            Actualizar_BD(NSE_, "SUMA", "HOGARES", Ciudad_, Mercado_, "HOGARES", "MENSUAL", Periodo, 
                                valor, int.Parse(BD_Zoho.sCabecera48Meses[i - 3].Substring(0, 4)));
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

        private void Add_Dapper(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, string _Periodo, double _Datos, int _Año)
        {
            using (IDbConnection dbDapper = new  SqlConnection(ConfigurationManager.ConnectionStrings["Zoho"].ConnectionString))
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
