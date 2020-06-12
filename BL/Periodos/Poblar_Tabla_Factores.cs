using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;

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
        
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string xAños, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2)
        {            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_FACTOR_AÑOS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
                //db_Zoho.AddInParameter(cmd_1, "_PER12M_1", DbType.String, _PER12M_1);
                //db_Zoho.AddInParameter(cmd_1, "_PER12M_2", DbType.String, _PER12M_2);
                //db_Zoho.AddInParameter(cmd_1, "_PER6M_1", DbType.String, _PER6M_1);
                //db_Zoho.AddInParameter(cmd_1, "_PER6M_2", DbType.String, _PER6M_2);
                //db_Zoho.AddInParameter(cmd_1, "_PER3M_1", DbType.String, _PER3M_1);
                //db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, _PER3M_2);
                //db_Zoho.AddInParameter(cmd_1, "_PER1M_1", DbType.String, _PER1M_1);
                //db_Zoho.AddInParameter(cmd_1, "_PER1M_2", DbType.String, _PER1M_2);
                //db_Zoho.AddInParameter(cmd_1, "_PERYTDM_1", DbType.String, _PERYTDM_1);
                //db_Zoho.AddInParameter(cmd_1, "_PERYTDM_2", DbType.String, _PERYTDM_2);
                int rows = 0;

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
                                switch (reader_1[i])
                                {
                                    case 2:
                                        valor_1 = 0;
                                        break;
                                    case 3:
                                        valor_2 = 0;
                                        break;
                                    case 4:
                                        valor_3 = 0;
                                        break;
                                }
                            }
                            else
                            {
                                switch (i)
                                {
                                    case 2:
                                        valor_1 = double.Parse(reader_1[i].ToString());
                                        break;
                                    case 3:
                                        valor_2 = double.Parse(reader_1[i].ToString());
                                        break;
                                    case 4:
                                        valor_3 = double.Parse(reader_1[i].ToString());
                                        break;
                                }
                                //valor_1 = double.Parse(reader_1[i].ToString()) == 2.0 ? 0 : 1;
                            }
                        }
                        rows++;
                        Actualizar_BD(V1, V2, valor_1, valor_2, valor_3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
                    }                    
                }
            }
        }          

        private void Actualizar_BD(int _V1, int _V2, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1, double _PER_12M_2, double _PER_6M_1, double _PER_6M_2, double _PER_3M_1, double _PER_3M_2, double _PER_1M_1, double _PER_1M_2, double _PER_YTD_1, double _PER_YTD_2)
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

    }
}