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
        double valor_1, valor_2, valor_3;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");
        
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string _PER_AÑO_0, string _PER_AÑO_1, string _PER_AÑO_2)
        {
            DbCommand cmdDelete;
            cmdDelete = db.GetSqlStringCommand("TRUNCATE TABLE BIP.dbo.FACTORES_HOGAR_PERIODOS");
            db.ExecuteNonQuery(cmdDelete);
            cmdDelete = db.GetSqlStringCommand("TRUNCATE TABLE BIP.dbo.BASE_RESULT_HOGAR_PERIODOS");
            db.ExecuteNonQuery(cmdDelete);

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_A_POBLAR_TABLA"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, _PER12M_1);
                db_Zoho.ExecuteNonQuery(cmd_1);
            }

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
                            Insertar_Registros(V1, V2, 0, 0, 0, valor_1, 0, 0, 0, 0, 0, 0, 0, 0, 0);
                        }
                        rows++;
                    }
                }
            }

            Update_Periodos(xCiudad, xCab, _PER12M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_12M_2]");
            Update_Periodos(xCiudad, xCab, _PER6M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_6M_1]");
            Update_Periodos(xCiudad, xCab, _PER6M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_6M_2]");
            Update_Periodos(xCiudad, xCab, _PER3M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_3M_1]");
            Update_Periodos(xCiudad, xCab, _PER3M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_3M_2]");
            Update_Periodos(xCiudad, xCab, _PER1M_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_1M_1]");
            Update_Periodos(xCiudad, xCab, _PER1M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_1M_2]");
            Update_Periodos(xCiudad, xCab, _PER1M_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_1M_2]");
            Update_Periodos(xCiudad, xCab, _PERYTDM_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_YTD_1]");
            Update_Periodos(xCiudad, xCab, _PERYTDM_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_YTD_2]");
            Update_Periodos(xCiudad, xCab, _PER_AÑO_0, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_0]");
            Update_Periodos(xCiudad, xCab, _PER_AÑO_1, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_1]");
            Update_Periodos(xCiudad, xCab, _PER_AÑO_2, "[PERIODOS].[_SP_FACTOR_UPDATE_PER_AÑO_2]");           
        }          

        private void Insertar_Registros(int _V1, int _V2, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1, double _PER_12M_2, double _PER_6M_1, double _PER_6M_2, double _PER_3M_1, double _PER_3M_2, double _PER_1M_1, double _PER_1M_2, double _PER_YTD_1, double _PER_YTD_2)
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
        private void Update_Factor_SP_FACTOR_12M(int _V1, int _V2, double valor, string procedimiento)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand(procedimiento))
            {
                db_Zoho.AddInParameter(cmd_1, "_V1", DbType.String, _V1);
                db_Zoho.AddInParameter(cmd_1, "_V2", DbType.String, _V2);
                db_Zoho.AddInParameter(cmd_1, "_VALOR", DbType.Double, valor);         
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }
        private void Update_Periodos(string xCiudad, string xCab, string xPer, string procedimiento)
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
    }
}