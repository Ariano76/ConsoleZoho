﻿using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;

namespace BL
{
    public class Share_Moneda_Local
    {
        private int[] Codigo_MARCA_VD = { 540, 5914, 1163, 504, 24, 1764, 318, 3206, 8019, 8420 }; // MARCAS PREDEFINIDAS VD 
        public string[] Codigo_MARCA_VR = new string[5];
        public string[] Codigo_MARCA_Nombre_VR = new string[5];

        private int[] Codigo_Belcorp = { 540, 1163, 5914 }; // GRUPO BELCORP
        private int[] Codigo_Loreal = { 5193, 696, 63, 1415, 3983, 1620, 3132, 176, 189, 1102, 2968, 6595, 1271, 293, 279, 282, 322, 583, 5434, 424, 423, 4225, 506, 1350, 8493, 513, 537, 209 }; // GRUPO LOREAL
        private int[] Codigo_Lauder = { 9266, 89, 1063, 1031, 174, 1299, 1588, 3181, 1009, 371, 987, 1398, 1128, 1068, 524 }; // GRUPO ESTEE LAUDER

        private int[] Codigo_NSE = { 1, 2, 3, 4, 5 }; // NSE
        private int[] Codigo_TIPOS = { 158, 161, 215, 202, 237, 226 }; // TIPOS 
        private int[] Codigo_Categoria = { 123, 124, 127 }; // Categorias      

        public double[] sShareValor;

        public string Codigo_Marcas_3M_Tipo, Codigo_Marcas_3M_Categoria, Codigo_Marcas_3M_Total_Cosmeticos;
        public string Codigo_Grupo_Belcorp_3M_Tipo, Codigo_Grupo_Loreal_3M_Tipo, Codigo_Grupo_Lauder_3M_Tipo;

        public double[,] periodo_Total_Mercado_Valor = new double[1, 13];  // TOTAL POR  VALORES
        public double[,] periodo_Total_Lima_Valor = new double[1, 13];  // LIMA POR  VALORES
        public double[,] periodo_Total_Ciudades_Valor = new double[1, 13];  // CIUDAD POR  VALORES
        public double[,] periodo_Temp_Valor = new double[1, 13];  // TEMPORAL POR  VALORES





        public double[,] sdata48Meses_x_Tipo_Total = new double[1, 48];  // TOTAL POR TIPO VALORES
        public double[,] sdata48Meses_x_Tipo_Marcas_Valores = new double[15, 49];  // 
        public double[,] sdata48Meses_x_Tipo_Marcas_Unidades = new double[15, 49];  // 
        public double[,] sdata48Meses_x_Tipo_Grupo_Marcas_Valores = new double[1, 48];  // 
        public double[,] sdata48Meses_x_Tipo_Grupo_Marcas_Unidades = new double[1, 48];  // 
        public double[,] sdata48Meses_x_Categoria_Marcas_Valores = new double[15, 49];  //  VALORES
        public double[,] sdata48Meses_x_Categoria_Marcas_Unidades = new double[15, 49];  //  UNIDADES
        public double[,] sdata48Meses_x_Categoria_Grupo_Marcas_Valores = new double[1, 48];  //  VALORES
        public double[,] sdata48Meses_x_Categoria_Grupo_Marcas_Unidades = new double[1, 48];  //  UNIDADES
        public double[,] sdata48Meses_x_Total_Marcas_Valores = new double[15, 49];  // TOTAL POR  VALORES
        public double[,] sdata48Meses_x_Total_Marcas_Unidades = new double[15, 49];  // TOTAL POR  UNIDADES
        public double[,] sdata48Meses_x_Total_Grupo_Marcas_Valores = new double[1, 48];  // TOTAL  COSMETICOS VALORES
        public double[,] sdata48Meses_x_Total_Grupo_Marcas_Unidades = new double[1, 48];  // TOTAL  COSMETICOS UNIDADES

        private readonly DateTime[] Periodos = new DateTime[7];
        string V1, Mercado, Periodo, xTipos_;
        public string resultadoBD;
        double valor_1;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        public void Recuperar_Marcas_Grupo_Belcorp(string xPeriodos, int xTipo)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Belcorp.GetLength(0); i++)
            {
                xIdMarca += Codigo_Belcorp[i].ToString() + ",";
            }
            Codigo_Grupo_Belcorp_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }
        public void Recuperar_Marcas_Grupo_Loreal(string xPeriodos, int xTipo)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Loreal.GetLength(0); i++)
            {
                xIdMarca += Codigo_Loreal[i].ToString() + ",";
            }
            Codigo_Grupo_Loreal_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }
        public void Recuperar_Marcas_Grupo_Lauder(string xPeriodos, int xTipo)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Lauder.GetLength(0); i++)
            {
                xIdMarca += Codigo_Lauder[i].ToString() + ",";
            }
            Codigo_Grupo_Lauder_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }

        public void Recuperar_Marcas_Top_5_Retail_x_Tipo(string xPeriodos, int xTipo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Tipo = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_PPU_MARCAS_TOP_5_TIPO_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_IDMONEDA", DbType.Int32, xMoneda);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = 0;
                    while (reader_1.Read())
                    {
                        Codigo_MARCA_VR[cont] = reader_1[cols].ToString();
                        Codigo_MARCA_Nombre_VR[cont] = reader_1[1].ToString();
                        xIdMarcaOtros += reader_1[cols].ToString() + ",";
                        cont++;
                    }
                }
                Codigo_Marcas_3M_Tipo += xIdMarcaOtros.Substring(0, xIdMarcaOtros.Length - 1);
            }
        }
        public void Recuperar_Marcas_Top_5_Retail_x_Categoria(string xPeriodos, int xCateg, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Categoria = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_PPU_MARCAS_TOP_5_CATEG_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xCateg);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_IDMONEDA", DbType.Int32, xMoneda);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = 0;
                    while (reader_1.Read())
                    {
                        Codigo_MARCA_VR[cont] = reader_1[cols].ToString();
                        Codigo_MARCA_Nombre_VR[cont] = reader_1[1].ToString();
                        xIdMarcaOtros += reader_1[cols].ToString() + ",";
                        cont++;
                    }
                }
                Codigo_Marcas_3M_Categoria += xIdMarcaOtros.Substring(0, xIdMarcaOtros.Length - 1);
            }
        }
        public void Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(string xPeriodos, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_PPU_MARCAS_TOP_5_TOTAL_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_IDMONEDA", DbType.Int32, xMoneda);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = 0;
                    while (reader_1.Read())
                    {
                        Codigo_MARCA_VR[cont] = reader_1[cols].ToString();
                        Codigo_MARCA_Nombre_VR[cont] = reader_1[1].ToString();
                        xIdMarcaOtros += reader_1[cols].ToString() + ",";
                        cont++;
                    }
                }
                Codigo_Marcas_3M_Total_Cosmeticos += xIdMarcaOtros.Substring(0, xIdMarcaOtros.Length - 1);
            }
        }

        #region TOTAL MERCADO
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1 , string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2 , string _PER1M_1, string _PER1M_2 , string _PERYTDM_1 , string _PERYTDM_2)
        {
            if (xCiudad == "1")
            {
                V1 = "Lima";              
            }
            else if (xCiudad == "1,2,5")
            {
                V1 = "Consolidado";
            }
            else
            {
                V1 = "Ciudades";
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);             
                db_Zoho.AddInParameter(cmd_1, "_PER12M_1", DbType.String, _PER12M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER12M_2", DbType.String, _PER12M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER6M_1", DbType.String, _PER6M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER6M_2", DbType.String, _PER6M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER3M_1", DbType.String, _PER3M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, _PER3M_2);
                db_Zoho.AddInParameter(cmd_1, "_PER1M_1", DbType.String, _PER1M_1);
                db_Zoho.AddInParameter(cmd_1, "_PER1M_2", DbType.String, _PER1M_2);
                db_Zoho.AddInParameter(cmd_1, "_PERYTDM_1", DbType.String, _PERYTDM_1);
                db_Zoho.AddInParameter(cmd_1, "_PERYTDM_2", DbType.String, _PERYTDM_2);
                int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
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

                            if (V1 == "Consolidado")
                            {
                                periodo_Total_Mercado_Valor[rows, i - 1] = valor_1;
                                periodo_Temp_Valor[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Valor[rows, i - 1] = valor_1;
                                periodo_Temp_Valor[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Valor[rows, i - 1] = valor_1;
                                periodo_Temp_Valor[rows, i - 1] = valor_1;
                            }
                            
                        }
                        rows++;
                    }
                }
                Actualizar_BD(V1, "Suma", "MONEDA LOCAL", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL PERCY", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1, "Suma", "MONEDA LOCAL MES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL PERCY", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12 , periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);

            }


            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_MARCAS_VALOR_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
            //    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Marcas_Valores[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }

            //    for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
            //    {
            //        V1 = Validar_marca(sdata48Meses_x_Total_Marcas_Valores[i, 0].ToString());
            //        for (int x = 1; x < sdata48Meses_x_Total_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
            //        {
            //            if (x < 10)
            //            {
            //                Periodo = "0" + (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
            //            }
            //            else
            //            {
            //                Periodo = (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
            //            }
            //            if (sdata48Meses_x_Total_Marcas_Valores[i, x] <= 0)
            //            {
            //                valor_1 = 0;
            //            }
            //            else
            //            {
            //                valor_1 = sdata48Meses_x_Total_Marcas_Valores[i, x] / sdata48Meses_x_Total_Marcas_Unidades[i, x];
            //            }
            //            Actualizar_BD(V1, "Suma", "PPU (ML)", "0. Consolidado", "0. Cosmeticos", "PPU (ML)", "MENSUAL", Periodo, valor_1,
            //                int.Parse(BD_Zoho.sCabecera48Meses[x - 1].Substring(0, 4)));
            //        }
            //    }
            //}

            //// GRUPO BELCORP
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Unidades[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }
            //}
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
            //    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Valores[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }

            //    for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
            //    {
            //        for (int x = 0; x < sdata48Meses_x_Total_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
            //        {
            //            if (x < 9)
            //            {
            //                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            else
            //            {
            //                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            if (sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] <= 0)
            //            {
            //                valor_1 = 0;
            //            }
            //            else
            //            {
            //                valor_1 = sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Total_Grupo_Marcas_Unidades[i, x];
            //            }
            //            Actualizar_BD("00.Belcorp", "Suma", "PPU (ML)", "0. Consolidado", "0. Cosmeticos", "PPU (ML)", "MENSUAL", Periodo, valor_1,
            //                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
            //        }
            //    }
            //}

            //// GRUPO LOREAL
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Unidades[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }
            //}
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
            //    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Valores[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }

            //    for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
            //    {
            //        for (int x = 0; x < sdata48Meses_x_Total_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
            //        {
            //            if (x < 9)
            //            {
            //                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            else
            //            {
            //                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            if (sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] <= 0)
            //            {
            //                valor_1 = 0;
            //            }
            //            else
            //            {
            //                valor_1 = sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Total_Grupo_Marcas_Unidades[i, x];
            //            }
            //            Actualizar_BD("50.Total L'Oreal", "Suma", "PPU (ML)", "0. Consolidado", "0. Cosmeticos", "PPU (ML)", "MENSUAL", Periodo, valor_1,
            //                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
            //        }
            //    }
            //}

            //// GRUPO LAUDER
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Unidades[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }
            //}
            //using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_COSMETICOS"))
            //{
            //    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
            //    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
            //    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
            //    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
            //    int rows = 0;

            //    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //    {
            //        int cols = reader_1.FieldCount;
            //        while (reader_1.Read())
            //        {
            //            for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                if (reader_1[i] == DBNull.Value)
            //                {
            //                    valor_1 = 0;
            //                }
            //                else
            //                {
            //                    valor_1 = double.Parse(reader_1[i].ToString());
            //                }
            //                sdata48Meses_x_Total_Grupo_Marcas_Valores[rows, i] = valor_1;
            //            }
            //            rows++;
            //        }
            //    }

            //    for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
            //    {
            //        for (int x = 0; x < sdata48Meses_x_Total_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
            //        {
            //            if (x < 9)
            //            {
            //                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            else
            //            {
            //                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
            //            }
            //            if (sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] <= 0)
            //            {
            //                valor_1 = 0;
            //            }
            //            else
            //            {
            //                valor_1 = sdata48Meses_x_Total_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Total_Grupo_Marcas_Unidades[i, x];
            //            }
            //            Actualizar_BD("51.Total Estee Lauder", "Suma", "PPU (ML)", "0. Consolidado", "0. Cosmeticos", "PPU (ML)", "MENSUAL", Periodo, valor_1,
            //                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
            //        }
            //    }
            //}
        }     
        //public void Leer_Ultimos_48_Meses_TIPO_TOTAL(string Cab, string xPeriodos, int xTipo, int xMoneda)
        //{
        //    xTipos_ = xTipo.ToString();
        //    switch (xTipo)
        //    {
        //        case 158:
        //            Mercado = "01. Colonia Femeninas";
        //            break;
        //        case 161:
        //            Mercado = "02. Colonia Masculinas";
        //            break;
        //        case 215:
        //            Mercado = "09. Humectante/Nutritiva Corporal";
        //            break;
        //        case 202:
        //            Mercado = "08. Nutritiva Revit. Facial";
        //            break;
        //        case 237:
        //            Mercado = "14. Roll-On";
        //            break;
        //        case 226:
        //            Mercado = "10. Shampoo Adultos";
        //            break;
        //    }

        //    // MARCAS PREDEFINIDAS VD + TOP 5 OTROS
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_MARCAS_UNIDAD_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_MARCAS_VALOR_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            if (sdata48Meses_x_Tipo_Marcas_Valores[i, 0].ToString() != "")
        //            {
        //                V1 = Validar_marca(sdata48Meses_x_Tipo_Marcas_Valores[i, 0].ToString());
        //            }

        //            for (int x = 1; x < sdata48Meses_x_Tipo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 10)
        //                {
        //                    Periodo = "0" + (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
        //                }
        //                else
        //                {
        //                    Periodo = (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
        //                }
        //                if (sdata48Meses_x_Tipo_Marcas_Valores[i, x] <= 0 ||
        //                    sdata48Meses_x_Tipo_Marcas_Valores[i, x].ToString() == null)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Tipo_Marcas_Valores[i, x] / sdata48Meses_x_Tipo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD(V1, "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x - 1].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO BELCORP
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Tipo_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("00.Belcorp", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO LOREAL
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Tipo_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("50.Total L'Oreal", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO LAUDER
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_TIPO"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Tipo_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Tipo_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Tipo_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Tipo_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("51.Total Estee Lauder", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }
        //}
        //public void Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(string Cab, string xPeriodos, int xCategoria, int xMoneda)
        //{
        //    xTipos_ = xCategoria.ToString();
        //    switch (xCategoria)
        //    {
        //        case 123:
        //            Mercado = "1. Fragancias";
        //            break;
        //        case 124:
        //            Mercado = "2. Maquillaje";
        //            break;
        //        case 127:
        //            Mercado = "5. Cuidado Personal";
        //            break;
        //    }

        //    // MARCAS PREDEFINIDAS VD + TOP 5 OTROS
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_MARCAS_UNIDAD_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_MARCAS_VALOR_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.String, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            V1 = Validar_marca(sdata48Meses_x_Categoria_Marcas_Valores[i, 0].ToString());
        //            for (int x = 1; x < sdata48Meses_x_Categoria_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 10)
        //                {
        //                    Periodo = "0" + (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
        //                }
        //                else
        //                {
        //                    Periodo = (x) + ". " + BD_Zoho.sCabecera48Meses[x - 1];
        //                }
        //                if (sdata48Meses_x_Categoria_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Categoria_Marcas_Valores[i, x] / sdata48Meses_x_Categoria_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD(V1, "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x - 1].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO BELCORP
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Categoria_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("00.Belcorp", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO LOREAL
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Categoria_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("50.Total L'Oreal", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }

        //    // GRUPO LAUDER
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_UNIDAD_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }
        //    }
        //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("marcas._SP_GRUPO_MARCAS_VALOR_TOTAL_CATEGORIA"))
        //    {
        //        db_Zoho.AddInParameter(cmd_1, "_CATEG", DbType.String, xTipos_);
        //        db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
        //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
        //        db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
        //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
        //        int rows = 0;

        //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
        //        {
        //            int cols = reader_1.FieldCount;
        //            while (reader_1.Read())
        //            {
        //                for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
        //                {
        //                    if (reader_1[i] == DBNull.Value)
        //                    {
        //                        valor_1 = 0;
        //                    }
        //                    else
        //                    {
        //                        valor_1 = double.Parse(reader_1[i].ToString());
        //                    }
        //                    sdata48Meses_x_Categoria_Grupo_Marcas_Valores[rows, i] = valor_1;
        //                }
        //                rows++;
        //            }
        //        }

        //        for (int i = 0; i < rows; i++) // LEYENDO LAS FILAS DEL ARRAY
        //        {
        //            for (int x = 0; x < sdata48Meses_x_Categoria_Grupo_Marcas_Valores.GetLength(1); x++) //LEYENDO LAS COLUMNAS
        //            {
        //                if (x < 9)
        //                {
        //                    Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                else
        //                {
        //                    Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
        //                }
        //                if (sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] <= 0)
        //                {
        //                    valor_1 = 0;
        //                }
        //                else
        //                {
        //                    valor_1 = sdata48Meses_x_Categoria_Grupo_Marcas_Valores[i, x] / sdata48Meses_x_Categoria_Grupo_Marcas_Unidades[i, x];
        //                }
        //                Actualizar_BD("51.Total Estee Lauder", "Suma", "PPU (ML)", "0. Consolidado", Mercado, "PPU (ML)", "MENSUAL", Periodo, valor_1,
        //                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
        //            }
        //        }
        //    }
        //}
 
        #endregion

        private string Validar_marca(string codigo)
        {
            string Marca;
            if ("540" == codigo)
            {
                Marca = "01.Esika";
            }
            else if ("5914" == codigo)
            {
                Marca = "02.Cyzone";
            }
            else if ("1163" == codigo)
            {
                Marca = "03.Lbel";
            }
            else if ("504" == codigo)
            {
                Marca = "04.Unique";
            }
            else if ("24" == codigo)
            {
                Marca = "05.Avon";
            }
            else if ("1764" == codigo)
            {
                Marca = "06.Natura";
            }
            else if ("8420" == codigo)
            {
                Marca = "07.Dupree";
            }
            else if ("318" == codigo)
            {
                Marca = "08.Mary Kay";
            }
            else if ("8019" == codigo)
            {
                Marca = "09.Hinode";
            }
            else if ("3206" == codigo)
            {
                Marca = "10.Oriflame";
            }
            else if (Codigo_MARCA_VR[0].ToString() == codigo)
            {
                Marca = "11." + System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Codigo_MARCA_Nombre_VR[0].ToLower());
            }
            else if (Codigo_MARCA_VR[1].ToString() == codigo)
            {
                Marca = "12." + System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Codigo_MARCA_Nombre_VR[1].ToLower());
            }
            else if (Codigo_MARCA_VR[2].ToString() == codigo)
            {
                Marca = "13." + System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Codigo_MARCA_Nombre_VR[2].ToLower());
            }
            else if (Codigo_MARCA_VR[3].ToString() == codigo)
            {
                Marca = "14." + System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Codigo_MARCA_Nombre_VR[3].ToLower());
            }
            else if (Codigo_MARCA_VR[4].ToString() == codigo)
            {
                Marca = "15." + System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Codigo_MARCA_Nombre_VR[4].ToLower());
            }
            else
            {
                Marca = "XXX";
            }

            return Marca;
        }

        private void Actualizar_BD(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1 , double _PER_12M_2 , double _PER_6M_1 , double _PER_6M_2 , double _PER_3M_1 , double _PER_3M_2 , double _PER_1M_1 , double _PER_1M_2 , double _PER_YTD_1 , double _PER_YTD_2 )
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_INSERTA_DATA_PERIODOS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_V1", DbType.String, _V1);
                db_Zoho.AddInParameter(cmd_1, "_V2", DbType.String, _V2);
                db_Zoho.AddInParameter(cmd_1, "_VARIABLE", DbType.String, _Variable);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, _Ciudad);
                db_Zoho.AddInParameter(cmd_1, "_MERCADO", DbType.String, _Mercado);
                db_Zoho.AddInParameter(cmd_1, "_UNIDAD", DbType.String, _Unidad);
                db_Zoho.AddInParameter(cmd_1, "_REPORTE", DbType.String, _Reporte);                
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
                if (_Variable == "MONEDA LOCAL MES")
                {
                    db_Zoho.AddInParameter(cmd_1, "_VAR_ANO", DbType.Double, 0);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_12M", DbType.Double, 0);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_6M", DbType.Double, 0);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_3M", DbType.Double, 0);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_1M", DbType.Double, 0);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_YTD", DbType.Double, 0);
                }
                else
                {
                    db_Zoho.AddInParameter(cmd_1, "_VAR_ANO", DbType.Double, (_ANO_2 / _ANO_1 -1) * 100);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_12M", DbType.Double, (_PER_12M_2 / _PER_12M_1 -1) * 100);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_6M", DbType.Double, (_PER_6M_2 / _PER_6M_1 -1) * 100);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_3M", DbType.Double, (_PER_3M_2 / _PER_3M_1 -1) * 100);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_1M", DbType.Double, (_PER_1M_2 / _PER_1M_1 -1) * 100);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_YTD", DbType.Double, (_PER_YTD_2 / _PER_YTD_1 -1) * 100);
                }
               
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }

    }
}