using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;

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

        private int[] Codigos_NSE = { 488, 489, 490, 491, 492 }; // NSE
        private int[] Codigo_Tipos = { 229,238,164,158,161,174,215,178,202,237,226,228,173,235,175 }; // TIPOS 
        string[,] Codigo_Nombres_Tipo = new string[15, 2] { {"229","12. Acondic.Adultos-Niños-Bb"},{"238","15. Aerosol"},{"164","03. Colonia Baño"},{"158","01. Colonia Femeninas"},{"161","02. Colonia Masculinas"},{"175","06. Delineador Ojos"},{"174","05. Embellecedor-Rimmel"},{"215","09. Humectante/Nutritiva Corporal"},{"178","07. Labiales"},{"202","08. Nutritiva Revit. Facial"},{"237","14. Roll-On"},{"226","10. Shampoo Adultos"},{"228","11. Shampoo Bebes"},{"173","04. Sombras"},{"235","13. Trat.Capilar Adultos/Niños"}};  // codigos y nombre de tipos
        string[,] Codigo_Nombres_NSE = new string[5, 2] {{ "488", "Alto" }, { "489", "Medio" }, { "490", "Medio Bajo" }, { "491", "Bajo" }, { "492", "Muy Bajo" }};  // Codigos y Nombre de NSE
        string[,] Codigo_Nombres_Categoria = new string[5, 2] { { "123", "1. Fragancias" }, { "124", "2. Maquillaje" }, { "125", "3. Tratamiento Facial" }, { "126", "4. Tratamiento Corporal" }, { "127", "5. Cuidado Personal" } };  // Codigos y Nombre de CATEGORIAS
        private int[] Codigo_Categoria_All = { 123, 124, 125, 126, 127 }; // Categorias   
        private int[] Codigo_Categoria = { 123, 124, 127 }; // Categorias que se miden mensualmente
        private int[] Codigo_Modalidad_Venta = { 87,88 }; // Modalidad de Venta 


        public double[] sShareValor;

        public string Codigo_Marcas_3M_Tipo, Codigo_Marcas_3M_Categoria, Codigo_Marcas_3M_Total_Cosmeticos;
        public string Codigo_Grupo_Belcorp_3M_Tipo, Codigo_Grupo_Loreal_3M_Tipo, Codigo_Grupo_Lauder_3M_Tipo;
        public string Codigo_cadena_NSE; 
        public string Codigo_cadena_Categoria;
        public string Codigo_cadena_Tipos, Codigo_cadena_Modalidad;

        public double[,] periodo_Total_Mercado_Valor = new double[1, 13];  // TOTAL POR  VALORES
        public double[,] periodo_Total_Lima_Valor = new double[1, 13];  // LIMA POR  VALORES
        public double[,] periodo_Total_Ciudades_Valor = new double[1, 13];  // CIUDAD POR  VALORES
        public double[,] periodo_Temp_Valor = new double[1, 13];  // TEMPORAL POR  VALORES

        public double[,] periodo_Total_Mercado_Tipo_Valor = new double[15, 14];  // LIMA POR  VALORES
        public double[,] periodo_Total_Lima_Tipo_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Ciudades_Tipo_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Temp = new double[15, 14];  // 

        public double[,] periodo_Total_Mercado_NSE_Valor = new double[15, 14];  // NSE
        public double[,] periodo_Total_Lima_NSE_Valor = new double[15, 14];  // NSE
        public double[,] periodo_Total_Ciudades_NSE_Valor = new double[15, 14];  // NSE 
        public double[,] periodo_Total_Mercado_Categoria_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Lima_Categoria_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Ciudades_Categoria_Valor = new double[15, 14];  //  





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
        string V1, V1_, V2, Ciudad_, Mercado, Variable, Variable_Promedio, Periodo, xTipos_;
        public string resultadoBD;
        double valor_1;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        public void Recuperar_Codigos_NSE()
        {
            string xNSE = "";
            for (int i = 0; i < Codigos_NSE.GetLength(0); i++)
            {
                xNSE += Codigos_NSE[i].ToString() + ",";
            }
            Codigo_cadena_NSE = xNSE.Substring(0, xNSE.Length - 1);
        }
        public void Recuperar_Codigos_Categoria()
        {
            string xNSE = "";
            for (int i = 0; i < Codigo_Categoria_All.GetLength(0); i++)
            {
                xNSE += Codigo_Categoria_All[i].ToString() + ",";
            }
            Codigo_cadena_Categoria = xNSE.Substring(0, xNSE.Length - 1);
        }
        public void Recuperar_Codigos_Tipos()
        {
            string xNSE = "";
            for (int i = 0; i < Codigo_Tipos.GetLength(0); i++)
            {
                xNSE += Codigo_Tipos[i].ToString() + ",";
            }
            Codigo_cadena_Tipos = xNSE.Substring(0, xNSE.Length - 1);
        }
        public void Recuperar_Codigos_Modalidad()
        {
            string xNSE = "";
            for (int i = 0; i < Codigo_Modalidad_Venta.GetLength(0); i++)
            {
                xNSE += Codigo_Modalidad_Venta[i].ToString() + ",";
            }
            Codigo_cadena_Modalidad = xNSE.Substring(0, xNSE.Length - 1);
        }
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

        // TOTAL MERCADO
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2)
        {
            if (xMoneda == 1)
            {
                Variable = "MONEDA LOCAL";
                Variable_Promedio = "MONEDA LOCAL MES";
                V2 = "";
            }
            else
            {
                Variable = "DOLARES";
                Variable_Promedio = "DOLARES MES";
                V2 = "";
            }

            if (xCiudad == "1")
            {
                V1 = "Lima";
                Ciudad_ = "1. Capital";
                V1_ = "Cosmeticos";
            }
            else if (xCiudad == "1,2,5")
            {
                V1 = "Consolidado";
                Ciudad_ = "0. Consolidado";
            }
            else
            {
                V1 = "Ciudades";
                Ciudad_ = "2. Ciudades";
                V1_ = "Cosmeticos";
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
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

                    Actualizar_BD(V1, "Suma", Variable, "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                    // PROMEDIO MENSUAL
                    Actualizar_BD(V1, "Suma", Variable_Promedio, "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);

                    if (xCiudad != "1,2,5") //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS
                    {
                        Actualizar_BD("Cosmeticos", "Suma", Variable, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                        // PROMEDIO MENSUAL
                        Actualizar_BD("Cosmeticos", "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                    }
                }
            }

            DateTime HoraStart;

            // CODIGO RECORRIDO POR NSE   
            HoraStart = DateTime.Now;
            double[,] periodo_Total_Temp = new double[15, 14];

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_NSE"))
            {                
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
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
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
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
                                periodo_Total_Mercado_NSE_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_NSE_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_NSE_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                    
                    if (rows == 4)
                    {
                        periodo_Total_Temp[rows + 1, 0] = 489; periodo_Total_Temp[rows + 1, 1] = 0; 
                        periodo_Total_Temp[rows + 1, 2] = 0; periodo_Total_Temp[rows + 1, 3] = 0; 
                        periodo_Total_Temp[rows + 1, 4] = 0; periodo_Total_Temp[rows + 1, 5] = 0; 
                        periodo_Total_Temp[rows + 1, 6] = 0; periodo_Total_Temp[rows + 1, 7] = 0; 
                        periodo_Total_Temp[rows + 1, 8] = 0; periodo_Total_Temp[rows + 1, 9] = 0; 
                        periodo_Total_Temp[rows + 1, 10] = 0; periodo_Total_Temp[rows + 1, 11] = 0; 
                        periodo_Total_Temp[rows + 1, 12] = 0; periodo_Total_Temp[rows + 1, 13] = 0;
                    }
                }
            }

            string CodigoNSE;
            for (int k = 0; k < periodo_Total_Temp.GetLength(0); k++) //FILAS
            {                
                CodigoNSE = periodo_Total_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        Mercado = Codigo_Nombres_NSE[i, 1];
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Total_Temp[k, 1], periodo_Total_Temp[k, 2], periodo_Total_Temp[k, 3], periodo_Total_Temp[k, 4], periodo_Total_Temp[k, 5], periodo_Total_Temp[k, 6], periodo_Total_Temp[k, 7], periodo_Total_Temp[k, 8], periodo_Total_Temp[k, 9], periodo_Total_Temp[k, 10], periodo_Total_Temp[k, 11], periodo_Total_Temp[0, 12], periodo_Total_Temp[k, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Total_Temp[k, 1] / 12, periodo_Total_Temp[k, 2] / 12, periodo_Total_Temp[k, 3] / 12, periodo_Total_Temp[k, 4] / 12, periodo_Total_Temp[k, 5] / 12, periodo_Total_Temp[k, 6] / 6, periodo_Total_Temp[k, 7] / 6, periodo_Total_Temp[k, 8] / 3, periodo_Total_Temp[k, 9] / 3, periodo_Total_Temp[k, 10] / 1, periodo_Total_Temp[k, 11] / 1, periodo_Total_Temp[k, 12], periodo_Total_Temp[k, 13]);
            }

            Tiempo_Proceso("PERIODOS SHARE " + V1 + " - " + V1_ + " - " + Ciudad_ + " . .", HoraStart);

            // CODIGO RECORRIDO POR CATEGORIAS   
            HoraStart = DateTime.Now;
            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_CATEGORIA"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
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
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
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
                                periodo_Total_Mercado_Categoria_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }                   
                }
            }

            string CodigoCategoria;
            for (int k = 0; k < periodo_Total_Temp.GetLength(0); k++) //FILAS
            {
                CodigoCategoria = periodo_Total_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Categoria[i, 1];
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Total_Temp[k, 1], periodo_Total_Temp[k, 2], periodo_Total_Temp[k, 3], periodo_Total_Temp[k, 4], periodo_Total_Temp[k, 5], periodo_Total_Temp[k, 6], periodo_Total_Temp[k, 7], periodo_Total_Temp[k, 8], periodo_Total_Temp[k, 9], periodo_Total_Temp[k, 10], periodo_Total_Temp[k, 11], periodo_Total_Temp[0, 12], periodo_Total_Temp[k, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Total_Temp[k, 1] / 12, periodo_Total_Temp[k, 2] / 12, periodo_Total_Temp[k, 3] / 12, periodo_Total_Temp[k, 4] / 12, periodo_Total_Temp[k, 5] / 12, periodo_Total_Temp[k, 6] / 6, periodo_Total_Temp[k, 7] / 6, periodo_Total_Temp[k, 8] / 3, periodo_Total_Temp[k, 9] / 3, periodo_Total_Temp[k, 10] / 1, periodo_Total_Temp[k, 11] / 1, periodo_Total_Temp[k, 12], periodo_Total_Temp[k, 13]);
            }

            Tiempo_Proceso("PERIODOS SHARE " + V1 + " - " + V1_ + " - " + Ciudad_ + " . .", HoraStart);







            //// CODIGO RECORRIDO POR CATEGORIAS  
            //HoraStart = DateTime.Now;
            //string CodigoCategoria;
            //for (int k = 0; k < Codigo_Categoria_All.GetLength(0); k++)
            //{
            //    CodigoCategoria = Codigo_Categoria_All[k].ToString();
            //    switch (Codigo_Categoria_All[k])
            //    {
            //        case 123:
            //            V1_ = "Fragancias";
            //            Mercado = "1. Fragancias";
            //            break;
            //        case 124:
            //            V1_ = "Maquillaje";
            //            Mercado = "2. Maquillaje";
            //            break;
            //        case 125:
            //            V1_ = "Tratamiento Facial";
            //            Mercado = "3. Tratamiento Facial";
            //            break;
            //        case 126:
            //            V1_ = "Tratamiento Corporal";
            //            Mercado = "4. Tratamiento Corporal";
            //            break;
            //        case 127:
            //            V1_ = "Cuidado Personal";
            //            Mercado = "5. Cuidado Personal";
            //            break;
            //    }

            //    using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MERCADO"))
            //    {
            //        db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, CodigoCategoria);
            //        db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
            //        db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
            //        db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
            //        db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
            //        db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
            //        db_Zoho.AddInParameter(cmd_1, "_PER12M_1", DbType.String, _PER12M_1);
            //        db_Zoho.AddInParameter(cmd_1, "_PER12M_2", DbType.String, _PER12M_2);
            //        db_Zoho.AddInParameter(cmd_1, "_PER6M_1", DbType.String, _PER6M_1);
            //        db_Zoho.AddInParameter(cmd_1, "_PER6M_2", DbType.String, _PER6M_2);
            //        db_Zoho.AddInParameter(cmd_1, "_PER3M_1", DbType.String, _PER3M_1);
            //        db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, _PER3M_2);
            //        db_Zoho.AddInParameter(cmd_1, "_PER1M_1", DbType.String, _PER1M_1);
            //        db_Zoho.AddInParameter(cmd_1, "_PER1M_2", DbType.String, _PER1M_2);
            //        db_Zoho.AddInParameter(cmd_1, "_PERYTDM_1", DbType.String, _PERYTDM_1);
            //        db_Zoho.AddInParameter(cmd_1, "_PERYTDM_2", DbType.String, _PERYTDM_2);
            //        int rows = 0;
            //        int columnas = 0;
            //        using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
            //        {
            //            int cols = reader_1.FieldCount;
            //            columnas = cols;
            //            while (reader_1.Read())
            //            {
            //                for (int i = 1; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //                {
            //                    if (reader_1[i] == DBNull.Value)
            //                    {
            //                        valor_1 = 0;
            //                    }
            //                    else
            //                    {
            //                        valor_1 = double.Parse(reader_1[i].ToString());
            //                    }

            //                    if (V1 == "Consolidado")
            //                    {
            //                        periodo_Total_Mercado_Valor[rows, i - 1] = valor_1;
            //                        periodo_Temp_Valor[rows, i - 1] = valor_1;
            //                    }
            //                    else if (V1 == "Lima")
            //                    {
            //                        periodo_Total_Lima_Valor[rows, i - 1] = valor_1;
            //                        periodo_Temp_Valor[rows, i - 1] = valor_1;
            //                    }
            //                    else
            //                    {
            //                        periodo_Total_Ciudades_Valor[rows, i - 1] = valor_1;
            //                        periodo_Temp_Valor[rows, i - 1] = valor_1;
            //                    }
            //                }
            //                rows++;
            //            }
            //        }
            //        if (rows < 1)
            //        {
            //            for (int i = 1; i < columnas; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
            //            {
            //                periodo_Temp_Valor[0, i - 1] = 0;
            //            }
            //        }
            //        Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
            //        // PROMEDIO MENSUAL
            //        Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
            //    }
            //}
            //Tiempo_Proceso("PERIODOS SHARE " + V1 + " - " + V1_ + " - " + Ciudad_ + " . .", HoraStart);

            // CODIGO RECORRIDO POR TIPOS      
            HoraStart = DateTime.Now;
            string CodigoTipo;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_TIPO_V1"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
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
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
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
                                periodo_Total_Mercado_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Total_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Total_Temp.GetLength(0); ii++) // FILAS
            {
                CodigoTipo = periodo_Total_Temp[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                    {
                        V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado = Codigo_Nombres_Tipo[i, 1];
                    }
                }
                
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Total_Temp[ii, 1], periodo_Total_Temp[ii, 2], periodo_Total_Temp[ii, 3], periodo_Total_Temp[ii, 4], periodo_Total_Temp[ii, 5], periodo_Total_Temp[ii, 6], periodo_Total_Temp[ii, 7], periodo_Total_Temp[ii, 8], periodo_Total_Temp[ii, 9], periodo_Total_Temp[ii, 10], periodo_Total_Temp[ii, 11], periodo_Total_Temp[ii, 12], periodo_Total_Temp[ii, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Total_Temp[ii, 1] / 12, periodo_Total_Temp[ii, 2] / 12, periodo_Total_Temp[ii, 3] / 12, periodo_Total_Temp[ii, 4] / 12, periodo_Total_Temp[ii, 5] / 12, periodo_Total_Temp[ii, 6] / 6, periodo_Total_Temp[ii, 7] / 6, periodo_Total_Temp[ii, 8] / 3, periodo_Total_Temp[ii, 9] / 3, periodo_Total_Temp[ii, 10] / 1, periodo_Total_Temp[ii, 11] / 1, periodo_Total_Temp[ii, 12], periodo_Total_Temp[ii, 13]);              
            }
            Tiempo_Proceso("PERIODOS SHARE " + V1 + " - " + V1_ + " - " + Ciudad_ + " . .", HoraStart);

            // CODIGO RECORRIDO POR MODALIDAD      
            HoraStart = DateTime.Now;
            string CodigoModalidad;
            for (int k = 0; k < Codigo_Modalidad_Venta.GetLength(0); k++)
            {
                CodigoModalidad = Codigo_Modalidad_Venta[k].ToString();
                switch (Codigo_Modalidad_Venta[k])
                {
                    case 87:
                        V1_ = "VD";
                        Mercado = "1. VD";
                        break;
                    case 88:
                        V1_ = "VR";
                        Mercado = "2. VR";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MODALIDAD", DbType.String, CodigoModalidad);
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
                    int columnas = 0;
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        columnas = cols;
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
                    if (rows < 1)
                    {
                        for (int i = 1; i < columnas; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            periodo_Temp_Valor[0, i - 1] = 0;
                        }
                    }
                    Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                    // PROMEDIO MENSUAL
                    Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                }
            }
            Tiempo_Proceso("PERIODOS SHARE " + V1 + " - " + V1_ + " - " + Ciudad_ + " . .", HoraStart);

        }

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

        private void Actualizar_BD(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1, double _PER_12M_2, double _PER_6M_1, double _PER_6M_2, double _PER_3M_1, double _PER_3M_2, double _PER_1M_1, double _PER_1M_2, double _PER_YTD_1, double _PER_YTD_2)
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

                if (_Variable == "MONEDA LOCAL MES" || _Variable == "DOLARES MES")
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
                    double Dato_1 = 0, Dato_2 = 0, Dato_3 = 0, Dato_4 = 0, Dato_5 = 0, Dato_6 = 0;
                    if (_ANO_1 != 0) 
                    {
                        Dato_1 = (_ANO_2 / _ANO_1 - 1) * 100; 
                    }
                    if (_PER_12M_1 != 0)
                    {
                        Dato_2 = (_PER_12M_2 / _PER_12M_1 - 1) * 100;
                    }
                    if (_PER_6M_1 != 0)
                    {
                        Dato_3 = (_PER_6M_2 / _PER_6M_1 - 1) * 100;
                    }
                    if (_PER_3M_1 != 0)
                    {
                        Dato_4 = (_PER_3M_2 / _PER_3M_1 - 1) * 100;
                    }
                    if (_PER_1M_1 != 0)
                    {
                        Dato_5 = (_PER_1M_2 / _PER_1M_1 - 1) * 100;
                    }
                    if (_PER_YTD_1 != 0)
                    {
                        Dato_6 = (_PER_YTD_2 / _PER_YTD_1 - 1) * 100;
                    }

                    db_Zoho.AddInParameter(cmd_1, "_VAR_ANO", DbType.Double, Dato_1);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_12M", DbType.Double, Dato_2);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_6M", DbType.Double, Dato_3);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_3M", DbType.Double, Dato_4);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_1M", DbType.Double, Dato_5);
                    db_Zoho.AddInParameter(cmd_1, "_VAR_YTD", DbType.Double, Dato_6);
                }
                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }
        static void Tiempo_Proceso(string Proceso, DateTime Inicio)
        {
            TimeSpan HoraFin;
            HoraFin = DateTime.Now - Inicio;
            Debug.Print($"El Proceso {Proceso} tardo {HoraFin} segundos.");
        }

    }
}