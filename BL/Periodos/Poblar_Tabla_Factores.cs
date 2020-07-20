using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;
using System.Diagnostics;

namespace BL
{
    public class Poblar_Tabla_Factores
    {

        private double[,] periodo_NSE_CAT_UU = new double[25, 14];  //  
        private double[,] periodo_NSE_CAT_VAL = new double[25, 14];  //  
        private double[,] periodo_NSE_TIPO_UU = new double[75, 14];  //  
        private double[,] periodo_NSE_TIPO_VAL = new double[75, 14];  //  
        private double[,] periodo_TIPO_MOD_UU = new double[30, 14];  //  
        private double[,] periodo_TIPO_MOD_VAL = new double[30, 14];  // 
        private double[,] periodo_CAT_MOD_UU = new double[10, 14];  //  
        private double[,] periodo_CAT_MOD_VAL = new double[10, 14];  //
        
        int V1, V2;        
        double valor_1;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        private void Crear_Tabla_Universos_Periodos(string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string _PER_AÑO_0, string _PER_AÑO_1, string _PER_AÑO_2)
        {

            using (DbCommand cmd_0 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIVERSOS_PERIODOS_CIUDADES"))
            {
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_0", DbType.String, _PER_AÑO_0);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_1", DbType.String, _PER_AÑO_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_2", DbType.String, _PER_AÑO_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_12M_1", DbType.String, _PER12M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_12M_2", DbType.String, _PER12M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_6M_1", DbType.String, _PER6M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_6M_2", DbType.String, _PER6M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_3M_1", DbType.String, _PER3M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_3M_2", DbType.String, _PER3M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_1M_1", DbType.String, _PER1M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_1M_2", DbType.String, _PER1M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_YTD_1", DbType.String, _PERYTDM_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_YTD_2", DbType.String, _PERYTDM_2);
                db_Zoho.ExecuteNonQuery(cmd_0);
            }

            using (DbCommand cmd_0 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIVERSOS_PERIODOS_PAIS"))
            {                
                db_Zoho.ExecuteNonQuery(cmd_0);
            }
        }
        public void Crear_Tabla_Datos_Periodos(string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string _PER_AÑO_0, string _PER_AÑO_1, string _PER_AÑO_2)
        {
           
            Insertar_Datos_Tabla_Periodos(_PER12M_1, "PERIODOS._SP_A_POBLAR_HOGAR_12M_1");
            Insertar_Datos_Tabla_Periodos(_PER12M_2, "PERIODOS._SP_A_POBLAR_HOGAR_12M_2");
            Insertar_Datos_Tabla_Periodos(_PER6M_1, "PERIODOS._SP_A_POBLAR_HOGAR_6M_1");
            Insertar_Datos_Tabla_Periodos(_PER6M_2, "PERIODOS._SP_A_POBLAR_HOGAR_6M_2");
            Insertar_Datos_Tabla_Periodos(_PER3M_1, "PERIODOS._SP_A_POBLAR_HOGAR_3M_1");
            Insertar_Datos_Tabla_Periodos(_PER3M_2, "PERIODOS._SP_A_POBLAR_HOGAR_3M_2");
            Insertar_Datos_Tabla_Periodos(_PER1M_1, "PERIODOS._SP_A_POBLAR_HOGAR_1M_1");
            Insertar_Datos_Tabla_Periodos(_PER1M_2, "PERIODOS._SP_A_POBLAR_HOGAR_1M_2");
            Insertar_Datos_Tabla_Periodos(_PERYTDM_1, "PERIODOS._SP_A_POBLAR_HOGAR_YTD_1");
            Insertar_Datos_Tabla_Periodos(_PERYTDM_2, "PERIODOS._SP_A_POBLAR_HOGAR_YTD_2");
            Insertar_Datos_Tabla_Periodos(_PER_AÑO_0, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_0");
            Insertar_Datos_Tabla_Periodos(_PER_AÑO_1, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_1");
            Insertar_Datos_Tabla_Periodos(_PER_AÑO_2, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_2");
        }

        public void Crear_Tabla_Factores(string xCiudad, string xCab, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string _PER_AÑO_0, string _PER_AÑO_1, string _PER_AÑO_2)
        {           
            using (DbCommand cmdDelete = db_Zoho.GetSqlStringCommand("TRUNCATE TABLE dbo.FACTORES_HOGAR_PERIODOS"))
            { db_Zoho.ExecuteNonQuery(cmdDelete); }

            using (DbCommand cmdDelete = db_Zoho.GetSqlStringCommand("TRUNCATE TABLE dbo.UNIVERSOS_HOGAR_PERIODOS"))
            { db_Zoho.ExecuteNonQuery(cmdDelete); }

            using (DbCommand cmd_0 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIVERSOS_PERIODOS"))
            {                
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_0", DbType.String, _PER_AÑO_0);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_1", DbType.String, _PER_AÑO_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_AÑO_2", DbType.String, _PER_AÑO_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_12M_1", DbType.String, _PER12M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_12M_2", DbType.String, _PER12M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_6M_1", DbType.String, _PER6M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_6M_2", DbType.String, _PER6M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_3M_1", DbType.String, _PER3M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_3M_2", DbType.String, _PER3M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_1M_1", DbType.String, _PER1M_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_1M_2", DbType.String, _PER1M_2);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_YTD_1", DbType.String, _PERYTDM_1);
                db_Zoho.AddInParameter(cmd_0, "_PERIODO_YTD_2", DbType.String, _PERYTDM_2);
                db_Zoho.ExecuteNonQuery(cmd_0);
            }
            Crear_Tabla_Universos_Periodos(_PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER1M_2, _PERYTDM_1, _PERYTDM_2, _PER_AÑO_0, _PER_AÑO_1, _PER_AÑO_2);

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_FACTOR_PERIODOS"))  
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODOS", DbType.String, _PER12M_1);
                int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        V2 = int.Parse(reader_1[0].ToString());
                        for (int i = 1; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            valor_1 = reader_1[i] == DBNull.Value ? 0 : double.Parse(reader_1[i].ToString());

                            switch (i)
                            {
                                case 1:
                                    V1 = 1;
                                    break;
                                case 2:
                                    V1 = 2;
                                    break;
                                case 3:
                                    V1 = 5;
                                    break;
                            }
                            Insertar_Factores(V1, V2, 0, 0, 0, valor_1, 0, 0, 0, 0, 0, 0, 0, 0, 0);
                        }
                        rows++;
                    }
                }
            }

            Update_Factor_Periodos(xCiudad, xCab, _PER12M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_12M_2]");
            Update_Factor_Periodos(xCiudad, xCab, _PER6M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_6M_1]");
            Update_Factor_Periodos(xCiudad, xCab, _PER6M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_6M_2]");
            Update_Factor_Periodos(xCiudad, xCab, _PER3M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_3M_1]");
            Update_Factor_Periodos(xCiudad, xCab, _PER3M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_3M_2]");
            Update_Factor_Periodos(xCiudad, xCab, _PER1M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_1M_1]");
            Update_Factor_Periodos(xCiudad, xCab, _PER1M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_1M_2]");
            Update_Factor_Periodos(xCiudad, xCab, _PERYTDM_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_YTD_1]");
            Update_Factor_Periodos(xCiudad, xCab, _PERYTDM_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_YTD_2]");
            Update_Factor_Periodos(xCiudad, xCab, _PER_AÑO_0, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_0]");
            Update_Factor_Periodos(xCiudad, xCab, _PER_AÑO_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_1]");
            Update_Factor_Periodos(xCiudad, xCab, _PER_AÑO_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_2]");

            Update_Base_Result_Hogar_Periodos();

            // ACUMULANDO LOS DATOS DE LOS DIFERENTES PERIODOS EN UNA SOLA TABLA.
            DbCommand cmdTruncate;
            cmdTruncate = db_Zoho.GetStoredProcCommand("PERIODOS._SP_BD_Resultados_Periodos");
            db_Zoho.ExecuteNonQuery(cmdTruncate);
        }

        private void Insertar_Factores(int _V1, int _V2, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1, double _PER_12M_2, double _PER_6M_1, double _PER_6M_2, double _PER_3M_1, double _PER_3M_2, double _PER_1M_1, double _PER_1M_2, double _PER_YTD_1, double _PER_YTD_2)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_FACTOR_INSERT"))
            {
                db_Zoho.AddInParameter(cmd_1, "_V1", DbType.String, _V1);
                db_Zoho.AddInParameter(cmd_1, "_V2", DbType.String, _V2);       
                db_Zoho.AddInParameter(cmd_1, "_ANO_0", DbType.Double, _ANO_0);
                db_Zoho.AddInParameter(cmd_1, "_ANO_1", DbType.Double, _ANO_1);
                db_Zoho.AddInParameter(cmd_1, "_ANO_2", DbType.Double, _ANO_2);
                db_Zoho.AddInParameter(cmd_1, "_PER_12M_1", DbType.Double, _PER_12M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER_12M_2", DbType.Double, _PER_12M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER_6M_1", DbType.Double, _PER_6M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER_6M_2", DbType.Double, _PER_6M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER_3M_1", DbType.Double, _PER_3M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER_3M_2", DbType.Double, _PER_3M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER_1M_1", DbType.Double, _PER_1M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER_1M_2", DbType.Double, _PER_1M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER_YTD_1", DbType.Double, _PER_YTD_1);
                db_Zoho.AddInParameter(cmd_1, "_PER_YTD_2", DbType.Double, _PER_YTD_2);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }
        private void Update_Factor_Periodos(string xCiudad, string xCab, string xPer, string procedimiento)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_FACTOR_PERIODOS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODOS", DbType.String, xPer);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        V2 = int.Parse(reader_1[0].ToString());
                        for (int i = 1; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }

                            switch (i)
                            {
                                case 1:
                                    V1 = 1;
                                    break;
                                case 2:
                                    V1 = 2;
                                    break;
                                case 3:
                                    V1 = 5;
                                    break;
                            }
                            Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, procedimiento);
                        }
                    }
                }
            }
        }
        private void Update_Base_Result_Hogar_Periodos()
        {

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_FACTOR_SELECT"))
            {
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        V1 = int.Parse(reader_1[0].ToString());
                        V2 = int.Parse(reader_1[1].ToString());
                        for (int i = 2; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            switch (i)
                            {
                                case 2: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_0_UPDATE_FACTOR"); break;
                                case 3: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_1_UPDATE_FACTOR"); break;
                                case 4: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_AÑO_2_UPDATE_FACTOR"); break;
                                case 5: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_12M_1_UPDATE_FACTOR"); break;
                                case 6: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_12M_2_UPDATE_FACTOR"); break;
                                case 7: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_6M_1_UPDATE_FACTOR"); break;
                                case 8: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_6M_2_UPDATE_FACTOR"); break;
                                case 9: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_3M_1_UPDATE_FACTOR"); break;
                                case 10: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_3M_2_UPDATE_FACTOR"); break;
                                case 11: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_1M_1_UPDATE_FACTOR"); break;
                                case 12: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_1M_2_UPDATE_FACTOR"); break;
                                case 13: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_YTD_1_UPDATE_FACTOR"); break;
                                case 14: Update_Factor_SP_FACTOR_12M(V1, V2, valor_1, "PERIODOS._SP_A_POBLAR_HOGAR_YTD_2_UPDATE_FACTOR"); break;
                            }
                        }
                    }
                }
            }
        }
        private void Insertar_Datos_Tabla_Periodos(string periodo, string procedimiento)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand(procedimiento))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, periodo);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }
        private void Update_Factor_SP_FACTOR_12M(int _V1, int _V2, double valor, string procedimiento)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand(procedimiento))
            {
                db_Zoho.AddInParameter(cmd_1, "_V1", DbType.Int32, _V1);
                db_Zoho.AddInParameter(cmd_1, "_V2", DbType.Int32, _V2);
                db_Zoho.AddInParameter(cmd_1, "_VALOR", DbType.Double, valor);         
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }        
        
    }
}