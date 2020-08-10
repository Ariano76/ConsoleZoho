using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;
using System.Runtime.InteropServices;

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
        int[] Cod_NSE_Temp = new int[5]; // NSE
        int[,] Cod_Tipo_Temp = new int[75,2];

        int[,] Codigo_Tipos_Categorias = new int[15, 2] { { 229, 127 }, { 238, 127 }, { 164, 123 }, { 158, 123 }, { 161, 123 }, { 175, 124 }, { 174, 124 }, { 215, 126 }, { 178, 124 }, { 202, 125 }, { 237, 127 }, { 226, 127 }, { 228, 127 }, { 173, 124 }, { 235, 127 } };  // codigos de tipos y categorias
        string[,] Codigo_Nombres_Tipo = new string[15, 2] { {"229","12. Acondic.Adultos-Niños-Bb"},{"238","15. Aerosol"},{"164","03. Colonia Baño"},{"158","01. Colonia Femeninas"},{"161","02. Colonia Masculinas"},{"175","06. Delineador Ojos"},{"174","05. Embellecedor-Rimmel"},{"215","09. Humectante/Nutritiva Corporal"},{"178","07. Labiales"},{"202","08. Nutritiva Revit. Facial"},{"237","14. Roll-On"},{"226","10. Shampoo Adultos"},{"228","11. Shampoo Bebes"},{"173","04. Sombras"},{"235","13. Trat.Capilar Adultos/Niños"}};  // codigos y nombre de tipos
        string[,] Codigo_Nombres_NSE = new string[5, 2] {{ "488", "Alto" }, { "489", "Medio" }, { "490", "Medio Bajo" }, { "491", "Bajo" }, { "492", "Muy Bajo" }};  // Codigos y Nombre de NSE
        string[,] Codigo_Nombres_Categoria = new string[5, 2] { { "123", "1. Fragancias" }, { "124", "2. Maquillaje" }, { "125", "3. Tratamiento Facial" }, { "126", "4. Tratamiento Corporal" }, { "127", "5. Cuidado Personal" } };  // Codigos y Nombre de Categoria
        string[,] Codigo_Nombres_Modalidad = new string[2, 2] {{"87", "1. VD" }, { "88", "2. VR" }}; // Codigos y Nombre de Modalidad
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
        public double[,] periodo_Total_Mercado_Unidad = new double[1, 13];  // TOTAL POR  UNIDADES
        public double[,] periodo_Total_Lima_Unidad = new double[1, 13];  // LIMA POR  UNIDADES
        public double[,] periodo_Total_Ciudades_Unidad = new double[1, 13];  // CIUDAD POR  UNIDADES
        public double[,] periodo_Total_Mercado_Hogar = new double[1, 13];  // TOTAL POR  HOGAR
        public double[,] periodo_Total_Lima_Hogar = new double[1, 13];  // LIMA POR  HOGAR
        public double[,] periodo_Total_Ciudades_Hogar = new double[1, 13];  // CIUDAD POR  HOGAR
        public double[,] periodo_Temp_Valor = new double[1, 13];  // TEMPORAL POR  VALORES
        public double[,] periodo_Temp_Unidad = new double[1, 13];  // TEMPORAL POR  UNIDAD
        public double[,] periodo_Temp = new double[1, 13];  // TEMPORAL POR  UNIDAD
        public double[,] periodo_Total_Temporal = new double[1, 13];  //


        public double[,] periodo_Total_Mercado_Tipo_Valor = new double[15, 14];  // LIMA POR  VALORES
        public double[,] periodo_Total_Lima_Tipo_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Ciudades_Tipo_Valor = new double[15, 14];  // 
        public double[,] periodo_Total_Mercado_Tipo_Unidad = new double[15, 14];  // LIMA POR  UNIDAD
        public double[,] periodo_Total_Lima_Tipo_Unidad = new double[15, 14];  // 
        public double[,] periodo_Total_Ciudades_Tipo_Unidad = new double[15, 14];  // 
                       
        public double[,] periodo_Tipo_Temp = new double[15, 14];  // 
        public double[,] periodo_Tipo_Temp_Unid = new double[15, 14];  // 
        public double[,] periodo_Total_Mercado_Tipo_Hogar_ = new double[15, 14];  // UNIDAD PROMEDIO POR HOGAR
        public double[,] periodo_Total_Lima_Tipo_Hogar_ = new double[15, 14];  // UNIDAD PROMEDIO POR HOGAR
        public double[,] periodo_Total_Ciudades_Tipo_Hogar_ = new double[15, 14];  // UNIDAD PROMEDIO POR HOGAR
        public double[,] periodo_Tipo_Temporal = new double[15, 14];  // 


        public double[,] periodo_Universos_NSE = new double[15, 15];  // UNIVERSOS        
        public double[,] periodo_Universos = new double[3, 15];  // UNIVERSOS  

        public double[,] periodo_Total_Mercado_TipoNSE_Hogar = new double[75, 15];  // 
        public double[,] periodo_Total_Lima_TipoNSE_Hogar = new double[75, 15];  // 
        public double[,] periodo_Total_Ciudades_TipoNSE_Hogar = new double[75, 15];  // 
        public double[,] periodo_Tipo_NSE_Temp = new double[75, 15];  // 
        
        public double[,] periodo_Total_Mercado_CateNSE_Hogar = new double[25, 15];  // 
        public double[,] periodo_Total_Lima_CateNSE_Hogar = new double[25, 15];  // 
        public double[,] periodo_Total_Ciudades_CateNSE_Hogar = new double[25, 15];  // 
        public double[,] periodo_Cate_NSE_Temp = new double[25, 15];  //
        
        public double[,] periodo_Total_Mercado_TipoMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Total_Lima_TipoMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Total_Ciudades_TipoMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Tipo_MOD_Temp = new double[30, 15];  // 
        

        public double[,] periodo_Total_Mercado_CateMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Total_Lima_CateMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Total_Ciudades_CateMOD_Hogar = new double[30, 15];  // 
        public double[,] periodo_Cate_MOD_Temp = new double[30, 15];  // 

        public double[,] periodo_Total_Mercado_NSE_Valor = new double[15, 14];  // NSE Valor
        public double[,] periodo_Total_Lima_NSE_Valor = new double[15, 14];  // NSE Valor
        public double[,] periodo_Total_Ciudades_NSE_Valor = new double[15, 14];  // NSE Valor
        public double[,] periodo_Total_Mercado_NSE_Unidad = new double[15, 14];  // NSE Unidad
        public double[,] periodo_Total_Lima_NSE_Unidad = new double[15, 14];  // NSE Unidad
        public double[,] periodo_Total_Ciudades_NSE_Unidad = new double[15, 14];  // NSE Unidad
        public double[,] periodo_Total_Mercado_NSE_Hogar = new double[5, 14];  // NSE Hogar
        public double[,] periodo_Total_Lima_NSE_Hogar = new double[5, 14];  // NSE Hogar
        public double[,] periodo_Total_Ciudades_NSE_Hogar = new double[5, 14];  // NSE Hogar
        public double[,] periodo_NSE_Temp = new double[5, 14];  // NSE 
        public double[,] periodo_NSE_Temp_Unidad = new double[5, 14];  // NSE 
        public double[,] periodo_NSE_Temporal = new double[5, 14];  // PENETRACION 


        public double[,] periodo_Total_Mercado_Categoria_Valor = new double[5, 14];  // 
        public double[,] periodo_Total_Lima_Categoria_Valor = new double[5, 14];  // 
        public double[,] periodo_Total_Ciudades_Categoria_Valor = new double[5, 14];  // 
        public double[,] periodo_Total_Mercado_Categoria_Unidad = new double[5, 14];  // Unidad
        public double[,] periodo_Total_Lima_Categoria_Unidad = new double[5, 14];  // Unidad
        public double[,] periodo_Total_Ciudades_Categoria_Unidad = new double[5, 14];  // Unidad
        public double[,] periodo_Total_Mercado_Categoria_Hogar = new double[5, 14];  // Hogar
        public double[,] periodo_Total_Lima_Categoria_Hogar = new double[5, 14];  // Hogar
        public double[,] periodo_Total_Ciudades_Categoria_Hogar = new double[5, 14];  // Hogar
        public double[,] periodo_Cat_Temp = new double[5, 14];
        public double[,] periodo_Cat_Temp_Unid = new double[5, 14];
        public double[,] periodo_Cat_Temporal = new double[5, 14];  // Hogar

        public double[,] periodo_Total_Mercado_Modalidad_Valor = new double[2, 14];  // 
        public double[,] periodo_Total_Lima_Modalidad_Valor = new double[2, 14];  // 
        public double[,] periodo_Total_Ciudades_Modalidad_Valor = new double[2, 14];  // 
        public double[,] periodo_Total_Mercado_Modalidad_Unidad = new double[2, 14];  // Unidad
        public double[,] periodo_Total_Lima_Modalidad_Unidad = new double[2, 14];  // Unidad
        public double[,] periodo_Total_Ciudades_Modalidad_Unidad = new double[2, 14];  // Unidad
        public double[,] periodo_Total_Mercado_Modalidad_Hogar = new double[2, 14];  // 
        public double[,] periodo_Total_Lima_Modalidad_Hogar = new double[2, 14];  // 
        public double[,] periodo_Total_Ciudades_Modalidad_Hogar = new double[2, 14];  // 
        public double[,] periodo_Modalidad_HOG_Temporal = new double[2, 14];  // HOGARES
        public double[,] periodo_Modalidad_Temporal = new double[2, 14];  // PENETRACION

        public double[,] periodo_Canal_Temp = new double[2, 14];  //  
        public double[,] periodo_Canal_Temp_Unid = new double[2, 14];  //  

        public double[,] periodo_NSE_CAT_UU = new double[25, 14];  //  
        public double[,] periodo_NSE_CAT_VAL = new double[25, 15];  //  
        public double[,] periodo_NSE_TIPO_UU = new double[75, 14];  //  
        public double[,] periodo_NSE_TIPO_VAL = new double[75, 15];  //  
        public double[,] periodo_TIPO_MOD_UU = new double[30, 14];  //  
        public double[,] periodo_TIPO_MOD_VAL = new double[30, 15];  // 
        public double[,] periodo_CAT_MOD_UU = new double[10, 14];  //  
        public double[,] periodo_CAT_MOD_VAL = new double[10, 15];  //

        public double[,] _Tipo_Canal_Mercado_Unidad = new double[30, 15];  //
        public double[,] _Tipo_Canal_Lima_Unidad = new double[30, 15];  //
        public double[,] _Tipo_Canal_Ciudad_Unidad = new double[30, 15];  //
        public double[,] _Tipo_Canal_Valor = new double[30, 15];  //


        private readonly DateTime[] Periodos = new DateTime[7];
        string V1, V1_, V1__, V2, V2_, Ciudad_, Mercado, Mercado_, Variable, VariableGM, Variable_Promedio, Unidad;
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

        // VALORES MONEDA LOCAL - DOLARES
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, [Optional] int NumMeses)
        {
            if (xMoneda == 1)
            {
                Variable = "MONEDA LOCAL";
                VariableGM = "GASTO MEDIO (ML)";
                Variable_Promedio = "MONEDA LOCAL MES";
                V2 = "";
                Unidad = "PPU (ML)";
            }
            else
            {
                Variable = "DOLARES";
                VariableGM = "GASTO MEDIO (DOL.)";
                Variable_Promedio = "DOLARES MES";
                V2 = "";
                Unidad = "PPU (DOL.)";
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
                    // PPU
                    Actualizar_BD(V1, "Suma", Unidad, "0. Consolidado", "0. Cosmeticos", Unidad, "MENSUAL", periodo_Temp_Valor[0, 0] / periodo_Temp_Unidad[0, 0], periodo_Temp_Valor[0, 1] / periodo_Temp_Unidad[0, 1], periodo_Temp_Valor[0, 2] / periodo_Temp_Unidad[0, 2], periodo_Temp_Valor[0, 3] / periodo_Temp_Unidad[0, 3], periodo_Temp_Valor[0, 4] / periodo_Temp_Unidad[0, 4], periodo_Temp_Valor[0, 5] / periodo_Temp_Unidad[0, 5], periodo_Temp_Valor[0, 6] / periodo_Temp_Unidad[0, 6], periodo_Temp_Valor[0, 7] / periodo_Temp_Unidad[0, 7], periodo_Temp_Valor[0, 8] / periodo_Temp_Unidad[0, 8], periodo_Temp_Valor[0, 9] / periodo_Temp_Unidad[0, 9], periodo_Temp_Valor[0, 10] / periodo_Temp_Unidad[0, 10], periodo_Temp_Valor[0, 11] / periodo_Temp_Unidad[0, 11], periodo_Temp_Valor[0, 12] / periodo_Temp_Unidad[0, 12]);

                    // PROMEDIO MENSUAL
                    Actualizar_BD(V1, "Suma", Variable_Promedio, "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11] / NumMeses, periodo_Temp_Valor[0, 12] /NumMeses);

                    // CALCULAR UNIDADES PROMEDIO (HOG.)"
                    if (V1 == "Consolidado")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Hogar.Clone();
                    }
                    else if (V1 == "Lima")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Lima_Hogar.Clone();
                    }
                    else
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Hogar.Clone();
                    }

                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = double.Parse(periodo_Total_Temporal[0, 0].ToString()) != 0 ? periodo_Temp_Valor[0, 0] / periodo_Total_Temporal[0, 0] : 0;
                    t2 = double.Parse(periodo_Total_Temporal[0, 1].ToString()) != 0 ? periodo_Temp_Valor[0, 1] / periodo_Total_Temporal[0, 1] : 0;
                    t3 = double.Parse(periodo_Total_Temporal[0, 2].ToString()) != 0 ? periodo_Temp_Valor[0, 2] / periodo_Total_Temporal[0, 2] : 0;
                    t4 = double.Parse(periodo_Total_Temporal[0, 3].ToString()) != 0 ? periodo_Temp_Valor[0, 3] / periodo_Total_Temporal[0, 3] : 0;
                    t5 = double.Parse(periodo_Total_Temporal[0, 4].ToString()) != 0 ? periodo_Temp_Valor[0, 4] / periodo_Total_Temporal[0, 4] : 0;
                    t6 = double.Parse(periodo_Total_Temporal[0, 5].ToString()) != 0 ? periodo_Temp_Valor[0, 5] / periodo_Total_Temporal[0, 5] : 0;
                    t7 = double.Parse(periodo_Total_Temporal[0, 6].ToString()) != 0 ? periodo_Temp_Valor[0, 6] / periodo_Total_Temporal[0, 6] : 0;
                    t8 = double.Parse(periodo_Total_Temporal[0, 7].ToString()) != 0 ? periodo_Temp_Valor[0, 7] / periodo_Total_Temporal[0, 7] : 0;
                    t9 = double.Parse(periodo_Total_Temporal[0, 8].ToString()) != 0 ? periodo_Temp_Valor[0, 8] / periodo_Total_Temporal[0, 8] : 0;
                    t10 = double.Parse(periodo_Total_Temporal[0, 9].ToString()) != 0 ? periodo_Temp_Valor[0, 9] / periodo_Total_Temporal[0, 9] : 0;
                    t11 = double.Parse(periodo_Total_Temporal[0, 10].ToString()) != 0 ? periodo_Temp_Valor[0, 10] / periodo_Total_Temporal[0, 10] : 0;
                    t12 = double.Parse(periodo_Total_Temporal[0, 11].ToString()) != 0 ? periodo_Temp_Valor[0, 11] / periodo_Total_Temporal[0, 11] : 0;
                    t13 = double.Parse(periodo_Total_Temporal[0, 12].ToString()) != 0 ? periodo_Temp_Valor[0, 12] / periodo_Total_Temporal[0, 12] : 0;

                    Actualizar_BD(V1, "Suma", VariableGM, "0. Consolidado", "0. Cosmeticos", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                    if (xCiudad != "1,2,5") //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS / CAPITAL Y CIUDAD
                    {
                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, "0. Cosmeticos", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD("Cosmeticos", "Suma", Variable, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0], periodo_Temp_Valor[0, 1], periodo_Temp_Valor[0, 2], periodo_Temp_Valor[0, 3], periodo_Temp_Valor[0, 4], periodo_Temp_Valor[0, 5], periodo_Temp_Valor[0, 6], periodo_Temp_Valor[0, 7], periodo_Temp_Valor[0, 8], periodo_Temp_Valor[0, 9], periodo_Temp_Valor[0, 10], periodo_Temp_Valor[0, 11], periodo_Temp_Valor[0, 12]);
                        //PPU
                        Actualizar_BD("Cosmeticos", "Suma", Unidad, Ciudad_, "0. Cosmeticos", Unidad, "MENSUAL", periodo_Temp_Valor[0, 0] / periodo_Temp_Unidad[0, 0], periodo_Temp_Valor[0, 1] / periodo_Temp_Unidad[0, 1], periodo_Temp_Valor[0, 2] / periodo_Temp_Unidad[0, 2], periodo_Temp_Valor[0, 3] / periodo_Temp_Unidad[0, 3], periodo_Temp_Valor[0, 4] / periodo_Temp_Unidad[0, 4], periodo_Temp_Valor[0, 5] / periodo_Temp_Unidad[0, 5], periodo_Temp_Valor[0, 6] / periodo_Temp_Unidad[0, 6], periodo_Temp_Valor[0, 7] / periodo_Temp_Unidad[0, 7], periodo_Temp_Valor[0, 8] / periodo_Temp_Unidad[0, 8], periodo_Temp_Valor[0, 9] / periodo_Temp_Unidad[0, 9], periodo_Temp_Valor[0, 10] / periodo_Temp_Unidad[0, 10], periodo_Temp_Valor[0, 11] / periodo_Temp_Unidad[0, 11], periodo_Temp_Valor[0, 12] / periodo_Temp_Unidad[0, 12]);
                        // PROMEDIO MENSUAL
                        Actualizar_BD("Cosmeticos", "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_Temp_Valor[0, 0] / 12, periodo_Temp_Valor[0, 1] / 12, periodo_Temp_Valor[0, 2] / 12, periodo_Temp_Valor[0, 3] / 12, periodo_Temp_Valor[0, 4] / 12, periodo_Temp_Valor[0, 5] / 6, periodo_Temp_Valor[0, 6] / 6, periodo_Temp_Valor[0, 7] / 3, periodo_Temp_Valor[0, 8] / 3, periodo_Temp_Valor[0, 9] / 1, periodo_Temp_Valor[0, 10] / 1, periodo_Temp_Valor[0, 11] / NumMeses, periodo_Temp_Valor[0, 12] / NumMeses);                        
                    }

                    if (xCiudad != "1,2,5" && xMoneda == 2)
                    {
                        //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        // CALCULAR VALOR SHARE (%)
                        t1 = periodo_Temp_Valor[0, 0] / periodo_Total_Mercado_Valor[0, 0] * 100;
                        t2 = periodo_Temp_Valor[0, 1] / periodo_Total_Mercado_Valor[0, 1] * 100;
                        t3 = periodo_Temp_Valor[0, 2] / periodo_Total_Mercado_Valor[0, 2] * 100;
                        t4 = periodo_Temp_Valor[0, 3] / periodo_Total_Mercado_Valor[0, 3] * 100;
                        t5 = periodo_Temp_Valor[0, 4] / periodo_Total_Mercado_Valor[0, 4] * 100;
                        t6 = periodo_Temp_Valor[0, 5] / periodo_Total_Mercado_Valor[0, 5] * 100;
                        t7 = periodo_Temp_Valor[0, 6] / periodo_Total_Mercado_Valor[0, 6] * 100;
                        t8 = periodo_Temp_Valor[0, 7] / periodo_Total_Mercado_Valor[0, 7] * 100;
                        t9 = periodo_Temp_Valor[0, 8] / periodo_Total_Mercado_Valor[0, 8] * 100;
                        t10 = periodo_Temp_Valor[0, 9] / periodo_Total_Mercado_Valor[0, 9] * 100;
                        t11 = periodo_Temp_Valor[0, 10] / periodo_Total_Mercado_Valor[0, 10] * 100;
                        t12 = periodo_Temp_Valor[0, 11] / periodo_Total_Mercado_Valor[0, 11] * 100;
                        t13 = periodo_Temp_Valor[0, 12] / periodo_Total_Mercado_Valor[0, 12] * 100;

                        Actualizar_BD_Penetracion(V1, "%", "SHARE DOLARES", "0. Consolidado", "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }

            // CODIGO RECORRIDO POR NSE              
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
                                periodo_NSE_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_NSE_Valor[rows, i - 1] = valor_1;
                                periodo_NSE_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_NSE_Valor[rows, i - 1] = valor_1;
                                periodo_NSE_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }

                    if (rows == 4)
                    {
                        periodo_NSE_Temp[rows , 0] = 489; periodo_NSE_Temp[rows , 1] = 0;
                        periodo_NSE_Temp[rows , 2] = 0; periodo_NSE_Temp[rows , 3] = 0;
                        periodo_NSE_Temp[rows , 4] = 0; periodo_NSE_Temp[rows , 5] = 0;
                        periodo_NSE_Temp[rows , 6] = 0; periodo_NSE_Temp[rows , 7] = 0;
                        periodo_NSE_Temp[rows , 8] = 0; periodo_NSE_Temp[rows , 9] = 0;
                        periodo_NSE_Temp[rows , 10] = 0; periodo_NSE_Temp[rows , 11] = 0;
                        periodo_NSE_Temp[rows , 12] = 0; periodo_NSE_Temp[rows , 13] = 0;
                    }
                }
            }
            string CodigoNSE;
            for (int k = 0; k < periodo_NSE_Temp.GetLength(0); k++) //FILAS
            {
                CodigoNSE = periodo_NSE_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        Mercado = Codigo_Nombres_NSE[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_NSE_Temp[k, 1], periodo_NSE_Temp[k, 2], periodo_NSE_Temp[k, 3], periodo_NSE_Temp[k, 4], periodo_NSE_Temp[k, 5], periodo_NSE_Temp[k, 6], periodo_NSE_Temp[k, 7], periodo_NSE_Temp[k, 8], periodo_NSE_Temp[k, 9], periodo_NSE_Temp[k, 10], periodo_NSE_Temp[k, 11], periodo_NSE_Temp[0, 12], periodo_NSE_Temp[k, 13]);
                //PPU
                double ppu_1, ppu_2, ppu_3, ppu_4, ppu_5, ppu_6, ppu_7, ppu_8, ppu_9, ppu_10, ppu_11, ppu_12, ppu_13;
                ppu_1 = (periodo_NSE_Temp_Unidad[k, 1] == 0 ? 0 : periodo_NSE_Temp[k, 1] / periodo_NSE_Temp_Unidad[k, 1]);
                ppu_2 = (periodo_NSE_Temp_Unidad[k, 2] == 0 ? 0 : periodo_NSE_Temp[k, 2] / periodo_NSE_Temp_Unidad[k, 2]);
                ppu_3 = (periodo_NSE_Temp_Unidad[k, 3] == 0 ? 0 : periodo_NSE_Temp[k, 3] / periodo_NSE_Temp_Unidad[k, 3]);
                ppu_4 = (periodo_NSE_Temp_Unidad[k, 4] == 0 ? 0 : periodo_NSE_Temp[k, 4] / periodo_NSE_Temp_Unidad[k, 4]);
                ppu_5 = (periodo_NSE_Temp_Unidad[k, 5] == 0 ? 0 : periodo_NSE_Temp[k, 5] / periodo_NSE_Temp_Unidad[k, 5]);
                ppu_6 = (periodo_NSE_Temp_Unidad[k, 6] == 0 ? 0 : periodo_NSE_Temp[k, 6] / periodo_NSE_Temp_Unidad[k, 6]);
                ppu_7 = (periodo_NSE_Temp_Unidad[k, 7] == 0 ? 0 : periodo_NSE_Temp[k, 7] / periodo_NSE_Temp_Unidad[k, 7]);
                ppu_8 = (periodo_NSE_Temp_Unidad[k, 8] == 0 ? 0 : periodo_NSE_Temp[k, 8] / periodo_NSE_Temp_Unidad[k, 8]);
                ppu_9 = (periodo_NSE_Temp_Unidad[k, 9] == 0 ? 0 : periodo_NSE_Temp[k, 9] / periodo_NSE_Temp_Unidad[k, 9]);
                ppu_10 = (periodo_NSE_Temp_Unidad[k, 10] == 0 ? 0 : periodo_NSE_Temp[k, 10] / periodo_NSE_Temp_Unidad[k, 10]);
                ppu_11 = (periodo_NSE_Temp_Unidad[k, 11] == 0 ? 0 : periodo_NSE_Temp[k, 11] / periodo_NSE_Temp_Unidad[k, 11]);
                ppu_12 = (periodo_NSE_Temp_Unidad[k, 12] == 0 ? 0 : periodo_NSE_Temp[k, 12] / periodo_NSE_Temp_Unidad[k, 12]);
                ppu_13 = (periodo_NSE_Temp_Unidad[k, 13] == 0 ? 0 : periodo_NSE_Temp[k, 13] / periodo_NSE_Temp_Unidad[k, 13]);
       
                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, "0. Cosmeticos", Unidad, "MENSUAL", ppu_1, ppu_2, ppu_3, ppu_4, ppu_5, ppu_6, ppu_7, ppu_8, ppu_9, ppu_10, ppu_11, ppu_12, ppu_13);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", periodo_NSE_Temp[k, 1] / 12, periodo_NSE_Temp[k, 2] / 12, periodo_NSE_Temp[k, 3] / 12, periodo_NSE_Temp[k, 4] / 12, periodo_NSE_Temp[k, 5] / 12, periodo_NSE_Temp[k, 6] / 6, periodo_NSE_Temp[k, 7] / 6, periodo_NSE_Temp[k, 8] / 3, periodo_NSE_Temp[k, 9] / 3, periodo_NSE_Temp[k, 10] / 1, periodo_NSE_Temp[k, 11] / 1, periodo_NSE_Temp[k, 12] / NumMeses, periodo_NSE_Temp[k, 13] / NumMeses);

                // CALCULAR UNIDADES SHARE (%)"
                if (xMoneda == 1)
                {
                    if (V1 == "Consolidado")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Valor.Clone();
                    }
                    else if (V1 == "Lima")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Lima_Valor.Clone();
                    }
                    else
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Valor.Clone();
                    }

                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_NSE_Temp[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                    t2 = double.Parse(periodo_Total_Temporal[0, 1].ToString()) != 0 ? periodo_NSE_Temp[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                    t3 = double.Parse(periodo_Total_Temporal[0, 2].ToString()) != 0 ? periodo_NSE_Temp[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                    t4 = double.Parse(periodo_Total_Temporal[0, 3].ToString()) != 0 ? periodo_NSE_Temp[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                    t5 = double.Parse(periodo_Total_Temporal[0, 4].ToString()) != 0 ? periodo_NSE_Temp[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                    t6 = double.Parse(periodo_Total_Temporal[0, 5].ToString()) != 0 ? periodo_NSE_Temp[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                    t7 = double.Parse(periodo_Total_Temporal[0, 6].ToString()) != 0 ? periodo_NSE_Temp[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                    t8 = double.Parse(periodo_Total_Temporal[0, 7].ToString()) != 0 ? periodo_NSE_Temp[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                    t9 = double.Parse(periodo_Total_Temporal[0, 8].ToString()) != 0 ? periodo_NSE_Temp[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                    t10 = double.Parse(periodo_Total_Temporal[0, 9].ToString()) != 0 ? periodo_NSE_Temp[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                    t11 = double.Parse(periodo_Total_Temporal[0, 10].ToString()) != 0 ? periodo_NSE_Temp[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                    t12 = double.Parse(periodo_Total_Temporal[0, 11].ToString()) != 0 ? periodo_NSE_Temp[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                    t13 = double.Parse(periodo_Total_Temporal[0, 12].ToString()) != 0 ? periodo_NSE_Temp[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                }

                // CALCULAR VALORES PROMEDIO (HOG.)"
                //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                if (V1 == "Consolidado")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Mercado_NSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Lima_NSE_Hogar.Clone();
                }
                else
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Ciudades_NSE_Hogar.Clone();
                }
                for (int i = 0; i < periodo_NSE_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoNSE == periodo_NSE_Temporal[i, 0].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_NSE_Temporal[i, 1].ToString()) != 0 ? periodo_NSE_Temp[k, 1] / periodo_NSE_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_NSE_Temporal[i, 2].ToString()) != 0 ? periodo_NSE_Temp[k, 2] / periodo_NSE_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_NSE_Temporal[i, 3].ToString()) != 0 ? periodo_NSE_Temp[k, 3] / periodo_NSE_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_NSE_Temporal[i, 4].ToString()) != 0 ? periodo_NSE_Temp[k, 4] / periodo_NSE_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_NSE_Temporal[i, 5].ToString()) != 0 ? periodo_NSE_Temp[k, 5] / periodo_NSE_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_NSE_Temporal[i, 6].ToString()) != 0 ? periodo_NSE_Temp[k, 6] / periodo_NSE_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_NSE_Temporal[i, 7].ToString()) != 0 ? periodo_NSE_Temp[k, 7] / periodo_NSE_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_NSE_Temporal[i, 8].ToString()) != 0 ? periodo_NSE_Temp[k, 8] / periodo_NSE_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_NSE_Temporal[i, 9].ToString()) != 0 ? periodo_NSE_Temp[k, 9] / periodo_NSE_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_NSE_Temporal[i, 10].ToString()) != 0 ? periodo_NSE_Temp[k, 10] / periodo_NSE_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_NSE_Temporal[i, 11].ToString()) != 0 ? periodo_NSE_Temp[k, 11] / periodo_NSE_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_NSE_Temporal[i, 12].ToString()) != 0 ? periodo_NSE_Temp[k, 12] / periodo_NSE_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_NSE_Temporal[i, 13].ToString()) != 0 ? periodo_NSE_Temp[k, 13] / periodo_NSE_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, "0. Cosmeticos", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIAS               
            string CodigoCategoria;
            double cat_1, cat_2, cat_3, cat_4, cat_5, cat_6, cat_7, cat_8, cat_9, cat_10, cat_11, cat_12, cat_13;
            //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
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
                                periodo_Cat_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_Valor[rows, i - 1] = valor_1;
                                periodo_Cat_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_Valor[rows, i - 1] = valor_1;
                                periodo_Cat_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }            
            for (int k = 0; k < periodo_Cat_Temp.GetLength(0); k++) //FILAS
            {
                CodigoCategoria = periodo_Cat_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); 
                        Mercado = Codigo_Nombres_Categoria[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Cat_Temp[k, 1], periodo_Cat_Temp[k, 2], periodo_Cat_Temp[k, 3], periodo_Cat_Temp[k, 4], periodo_Cat_Temp[k, 5], periodo_Cat_Temp[k, 6], periodo_Cat_Temp[k, 7], periodo_Cat_Temp[k, 8], periodo_Cat_Temp[k, 9], periodo_Cat_Temp[k, 10], periodo_Cat_Temp[k, 11], periodo_Cat_Temp[0, 12], periodo_Cat_Temp[k, 13]);
                //PPU
                
                cat_1 = (periodo_Cat_Temp_Unid[k, 1] == 0 ? 0 : periodo_Cat_Temp[k, 1] / periodo_Cat_Temp_Unid[k, 1]);
                cat_2 = (periodo_Cat_Temp_Unid[k, 2] == 0 ? 0 : periodo_Cat_Temp[k, 2] / periodo_Cat_Temp_Unid[k, 2]);
                cat_3 = (periodo_Cat_Temp_Unid[k, 3] == 0 ? 0 : periodo_Cat_Temp[k, 3] / periodo_Cat_Temp_Unid[k, 3]);
                cat_4 = (periodo_Cat_Temp_Unid[k, 4] == 0 ? 0 : periodo_Cat_Temp[k, 4] / periodo_Cat_Temp_Unid[k, 4]);
                cat_5 = (periodo_Cat_Temp_Unid[k, 5] == 0 ? 0 : periodo_Cat_Temp[k, 5] / periodo_Cat_Temp_Unid[k, 5]);
                cat_6 = (periodo_Cat_Temp_Unid[k, 6] == 0 ? 0 : periodo_Cat_Temp[k, 6] / periodo_Cat_Temp_Unid[k, 6]);
                cat_7 = (periodo_Cat_Temp_Unid[k, 7] == 0 ? 0 : periodo_Cat_Temp[k, 7] / periodo_Cat_Temp_Unid[k, 7]);
                cat_8 = (periodo_Cat_Temp_Unid[k, 8] == 0 ? 0 : periodo_Cat_Temp[k, 8] / periodo_Cat_Temp_Unid[k, 8]);
                cat_9 = (periodo_Cat_Temp_Unid[k, 9] == 0 ? 0 : periodo_Cat_Temp[k, 9] / periodo_Cat_Temp_Unid[k, 9]);
                cat_10 = (periodo_Cat_Temp_Unid[k, 10] == 0 ? 0 : periodo_Cat_Temp[k, 10] / periodo_Cat_Temp_Unid[k, 10]);
                cat_11 = (periodo_Cat_Temp_Unid[k, 11] == 0 ? 0 : periodo_Cat_Temp[k, 11] / periodo_Cat_Temp_Unid[k, 11]);
                cat_12 = (periodo_Cat_Temp_Unid[k, 12] == 0 ? 0 : periodo_Cat_Temp[k, 12] / periodo_Cat_Temp_Unid[k, 12]);
                cat_13 = (periodo_Cat_Temp_Unid[k, 13] == 0 ? 0 : periodo_Cat_Temp[k, 13] / periodo_Cat_Temp_Unid[k, 13]);

                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", cat_1, cat_2, cat_3, cat_4, cat_5, cat_6, cat_7, cat_8, cat_9, cat_10, cat_11, cat_12, cat_13);
                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, "0. Cosmeticos2", Unidad, "MENSUAL", cat_1, cat_2, cat_3, cat_4, cat_5, cat_6, cat_7, cat_8, cat_9, cat_10, cat_11, cat_12, cat_13);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Cat_Temp[k, 1] / 12, periodo_Cat_Temp[k, 2] / 12, periodo_Cat_Temp[k, 3] / 12, periodo_Cat_Temp[k, 4] / 12, periodo_Cat_Temp[k, 5] / 12, periodo_Cat_Temp[k, 6] / 6, periodo_Cat_Temp[k, 7] / 6, periodo_Cat_Temp[k, 8] / 3, periodo_Cat_Temp[k, 9] / 3, periodo_Cat_Temp[k, 10] / 1, periodo_Cat_Temp[k, 11] / 1, periodo_Cat_Temp[k, 12] / NumMeses, periodo_Cat_Temp[k, 13] / NumMeses);

                // CALCULAR UNIDADES SHARE (%)"

                if (xMoneda == 1)
                {
                    if (V1 == "Consolidado")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Valor.Clone();
                    }
                    else if (V1 == "Lima")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Lima_Valor.Clone();
                    }
                    else
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Valor.Clone();
                    }

                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Cat_Temp[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                    t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Cat_Temp[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                    t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Cat_Temp[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                    t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Cat_Temp[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                    t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Cat_Temp[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                    t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Cat_Temp[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                    t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Cat_Temp[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                    t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Cat_Temp[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                    t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Cat_Temp[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                    t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Cat_Temp[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                    t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Cat_Temp[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                    t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Cat_Temp[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                    t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Cat_Temp[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                }

                // CALCULAR UNIDADES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Mercado_Categoria_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Lima_Categoria_Hogar.Clone();
                }
                else
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Ciudades_Categoria_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Cat_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoCategoria == periodo_Cat_Temporal[i, 0].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_Cat_Temporal[i, 1].ToString()) != 0 ? periodo_Cat_Temp[k, 1] / periodo_Cat_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_Cat_Temporal[i, 2].ToString()) != 0 ? periodo_Cat_Temp[k, 2] / periodo_Cat_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_Cat_Temporal[i, 3].ToString()) != 0 ? periodo_Cat_Temp[k, 3] / periodo_Cat_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_Cat_Temporal[i, 4].ToString()) != 0 ? periodo_Cat_Temp[k, 4] / periodo_Cat_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_Cat_Temporal[i, 5].ToString()) != 0 ? periodo_Cat_Temp[k, 5] / periodo_Cat_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_Cat_Temporal[i, 6].ToString()) != 0 ? periodo_Cat_Temp[k, 6] / periodo_Cat_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_Cat_Temporal[i, 7].ToString()) != 0 ? periodo_Cat_Temp[k, 7] / periodo_Cat_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_Cat_Temporal[i, 8].ToString()) != 0 ? periodo_Cat_Temp[k, 8] / periodo_Cat_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_Cat_Temporal[i, 9].ToString()) != 0 ? periodo_Cat_Temp[k, 9] / periodo_Cat_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_Cat_Temporal[i, 10].ToString()) != 0 ? periodo_Cat_Temp[k, 10] / periodo_Cat_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_Cat_Temporal[i, 11].ToString()) != 0 ? periodo_Cat_Temp[k, 11] / periodo_Cat_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_Cat_Temporal[i, 12].ToString()) != 0 ? periodo_Cat_Temp[k, 12] / periodo_Cat_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_Cat_Temporal[i, 13].ToString()) != 0 ? periodo_Cat_Temp[k, 13] / periodo_Cat_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, "0. Cosmeticos2", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado, VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPOS      
            string CodigoTipo;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO"))
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
                            periodo_Tipo_Temp_Unid[rows, i - 1] = valor_1;                           
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_TIPO"))
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

                            if (V1 == "Consolidado") {
                                periodo_Total_Mercado_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima") {
                                periodo_Total_Lima_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp[rows, i - 1] = valor_1;
                            }
                            else {
                                periodo_Total_Ciudades_Tipo_Valor[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int k = 0; k < periodo_Tipo_Temp.GetLength(0); k++) // FILAS
            {
                CodigoTipo = periodo_Tipo_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0]) {
                        V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado = Codigo_Nombres_Tipo[i, 1];
                        break;
                    }
                }

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Tipo_Temp[k, 1], periodo_Tipo_Temp[k, 2], periodo_Tipo_Temp[k, 3], periodo_Tipo_Temp[k, 4], periodo_Tipo_Temp[k, 5], periodo_Tipo_Temp[k, 6], periodo_Tipo_Temp[k, 7], periodo_Tipo_Temp[k, 8], periodo_Tipo_Temp[k, 9], periodo_Tipo_Temp[k, 10], periodo_Tipo_Temp[k, 11], periodo_Tipo_Temp[k, 12], periodo_Tipo_Temp[k, 13]);
                // PPU periodo_Total_Mercado_Tipo_Unidad
                double tip_1, tip_2, tip_3, tip_4, tip_5, tip_6, tip_7, tip_8, tip_9, tip_10, tip_11, tip_12, tip_13;
                tip_1 = (periodo_Tipo_Temp_Unid[k, 1] == 0 ? 0 : periodo_Tipo_Temp[k, 1] / periodo_Tipo_Temp_Unid[k, 1]);
                tip_2 = (periodo_Tipo_Temp_Unid[k, 2] == 0 ? 0 : periodo_Tipo_Temp[k, 2] / periodo_Tipo_Temp_Unid[k, 2]);
                tip_3 = (periodo_Tipo_Temp_Unid[k, 3] == 0 ? 0 : periodo_Tipo_Temp[k, 3] / periodo_Tipo_Temp_Unid[k, 3]);
                tip_4 = (periodo_Tipo_Temp_Unid[k, 4] == 0 ? 0 : periodo_Tipo_Temp[k, 4] / periodo_Tipo_Temp_Unid[k, 4]);
                tip_5 = (periodo_Tipo_Temp_Unid[k, 5] == 0 ? 0 : periodo_Tipo_Temp[k, 5] / periodo_Tipo_Temp_Unid[k, 5]);
                tip_6 = (periodo_Tipo_Temp_Unid[k, 6] == 0 ? 0 : periodo_Tipo_Temp[k, 6] / periodo_Tipo_Temp_Unid[k, 6]);
                tip_7 = (periodo_Tipo_Temp_Unid[k, 7] == 0 ? 0 : periodo_Tipo_Temp[k, 7] / periodo_Tipo_Temp_Unid[k, 7]);
                tip_8 = (periodo_Tipo_Temp_Unid[k, 8] == 0 ? 0 : periodo_Tipo_Temp[k, 8] / periodo_Tipo_Temp_Unid[k, 8]);
                tip_9 = (periodo_Tipo_Temp_Unid[k, 9] == 0 ? 0 : periodo_Tipo_Temp[k, 9] / periodo_Tipo_Temp_Unid[k, 9]);
                tip_10 = (periodo_Tipo_Temp_Unid[k, 10] == 0 ? 0 : periodo_Tipo_Temp[k, 10] / periodo_Tipo_Temp_Unid[k, 10]);
                tip_11 = (periodo_Tipo_Temp_Unid[k, 11] == 0 ? 0 : periodo_Tipo_Temp[k, 11] / periodo_Tipo_Temp_Unid[k, 11]);
                tip_12 = (periodo_Tipo_Temp_Unid[k, 12] == 0 ? 0 : periodo_Tipo_Temp[k, 12] / periodo_Tipo_Temp_Unid[k, 12]);
                tip_13 = (periodo_Tipo_Temp_Unid[k, 13] == 0 ? 0 : periodo_Tipo_Temp[k, 13] / periodo_Tipo_Temp_Unid[k, 13]);
                
                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", tip_1, tip_2, tip_3, tip_4, tip_5, tip_6, tip_7, tip_8, tip_9, tip_10, tip_11, tip_12, tip_13);

                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, "0. Cosmeticos2", Unidad, "MENSUAL", tip_1, tip_2, tip_3, tip_4, tip_5, tip_6, tip_7, tip_8, tip_9, tip_10, tip_11, tip_12, tip_13);

                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Tipo_Temp[k, 1] / 12, periodo_Tipo_Temp[k, 2] / 12, periodo_Tipo_Temp[k, 3] / 12, periodo_Tipo_Temp[k, 4] / 12, periodo_Tipo_Temp[k, 5] / 12, periodo_Tipo_Temp[k, 6] / 6, periodo_Tipo_Temp[k, 7] / 6, periodo_Tipo_Temp[k, 8] / 3, periodo_Tipo_Temp[k, 9] / 3, periodo_Tipo_Temp[k, 10] / 1, periodo_Tipo_Temp[k, 11] / 1, periodo_Tipo_Temp[k, 12] / NumMeses, periodo_Tipo_Temp[k, 13] / NumMeses);

                // CALCULAR VALORES SHARE (%)"
                if (xMoneda == 1)
                {
                    if (V1 == "Consolidado")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Valor.Clone();
                    }
                    else if (V1 == "Lima")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Lima_Valor.Clone();
                    }
                    else
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Valor.Clone();
                    }

                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Tipo_Temp[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                    t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Tipo_Temp[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                    t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Tipo_Temp[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                    t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Tipo_Temp[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                    t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Tipo_Temp[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                    t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Tipo_Temp[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                    t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Tipo_Temp[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                    t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Tipo_Temp[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                    t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Tipo_Temp[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                    t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Tipo_Temp[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                    t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Tipo_Temp[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                    t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Tipo_Temp[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                    t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Tipo_Temp[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                }

                // CALCULAR VALORES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Mercado_Tipo_Hogar_.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Lima_Tipo_Hogar_.Clone();
                }
                else
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Ciudades_Tipo_Hogar_.Clone();
                }
                
                for (int i = 0; i < periodo_Tipo_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    if (CodigoTipo == periodo_Tipo_Temporal[i, 0].ToString())
                    {                        
                        t1 = double.Parse(periodo_Tipo_Temporal[i, 1].ToString()) != 0 ? periodo_Tipo_Temp[k, 1] / periodo_Tipo_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_Tipo_Temporal[i, 2].ToString()) != 0 ? periodo_Tipo_Temp[k, 2] / periodo_Tipo_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_Tipo_Temporal[i, 3].ToString()) != 0 ? periodo_Tipo_Temp[k, 3] / periodo_Tipo_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_Tipo_Temporal[i, 4].ToString()) != 0 ? periodo_Tipo_Temp[k, 4] / periodo_Tipo_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_Tipo_Temporal[i, 5].ToString()) != 0 ? periodo_Tipo_Temp[k, 5] / periodo_Tipo_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_Tipo_Temporal[i, 6].ToString()) != 0 ? periodo_Tipo_Temp[k, 6] / periodo_Tipo_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_Tipo_Temporal[i, 7].ToString()) != 0 ? periodo_Tipo_Temp[k, 7] / periodo_Tipo_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_Tipo_Temporal[i, 8].ToString()) != 0 ? periodo_Tipo_Temp[k, 8] / periodo_Tipo_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_Tipo_Temporal[i, 9].ToString()) != 0 ? periodo_Tipo_Temp[k, 9] / periodo_Tipo_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_Tipo_Temporal[i, 10].ToString()) != 0 ? periodo_Tipo_Temp[k, 10] / periodo_Tipo_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_Tipo_Temporal[i, 11].ToString()) != 0 ? periodo_Tipo_Temp[k, 11] / periodo_Tipo_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_Tipo_Temporal[i, 12].ToString()) != 0 ? periodo_Tipo_Temp[k, 12] / periodo_Tipo_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_Tipo_Temporal[i, 13].ToString()) != 0 ? periodo_Tipo_Temp[k, 13] / periodo_Tipo_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, "0. Cosmeticos2", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado, VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR MODALIDAD      
            string CodigoModalidad;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MODALIDAD", DbType.String, Codigo_cadena_Modalidad);
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
                                periodo_Total_Mercado_Modalidad_Valor[rows, i - 1] = valor_1;
                                periodo_Canal_Temp[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Modalidad_Valor[rows, i - 1] = valor_1;
                                periodo_Canal_Temp[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Modalidad_Valor[rows, i - 1] = valor_1;
                                periodo_Canal_Temp[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int k = 0; k < periodo_Canal_Temp.GetLength(0); k++) //FILAS
            {
                CodigoModalidad = periodo_Canal_Temp[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                    {
                        V1_ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); 
                        Mercado = Codigo_Nombres_Modalidad[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Canal_Temp[k, 1], periodo_Canal_Temp[k, 2], periodo_Canal_Temp[k, 3], periodo_Canal_Temp[k, 4], periodo_Canal_Temp[k, 5], periodo_Canal_Temp[k, 6], periodo_Canal_Temp[k, 7], periodo_Canal_Temp[k, 8], periodo_Canal_Temp[k, 9], periodo_Canal_Temp[k, 10], periodo_Canal_Temp[k, 11], periodo_Canal_Temp[0, 12], periodo_Canal_Temp[k, 13]);
                //PPU
                double mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13;
                mod_1 = (periodo_Canal_Temp_Unid[k, 1] == 0 ? 0 : periodo_Canal_Temp[k, 1] / periodo_Canal_Temp_Unid[k, 1]);
                mod_2 = (periodo_Canal_Temp_Unid[k, 2] == 0 ? 0 : periodo_Canal_Temp[k, 2] / periodo_Canal_Temp_Unid[k, 2]);
                mod_3 = (periodo_Canal_Temp_Unid[k, 3] == 0 ? 0 : periodo_Canal_Temp[k, 3] / periodo_Canal_Temp_Unid[k, 3]);
                mod_4 = (periodo_Canal_Temp_Unid[k, 4] == 0 ? 0 : periodo_Canal_Temp[k, 4] / periodo_Canal_Temp_Unid[k, 4]);
                mod_5 = (periodo_Canal_Temp_Unid[k, 5] == 0 ? 0 : periodo_Canal_Temp[k, 5] / periodo_Canal_Temp_Unid[k, 5]);
                mod_6 = (periodo_Canal_Temp_Unid[k, 6] == 0 ? 0 : periodo_Canal_Temp[k, 6] / periodo_Canal_Temp_Unid[k, 6]);
                mod_7 = (periodo_Canal_Temp_Unid[k, 7] == 0 ? 0 : periodo_Canal_Temp[k, 7] / periodo_Canal_Temp_Unid[k, 7]);
                mod_8 = (periodo_Canal_Temp_Unid[k, 8] == 0 ? 0 : periodo_Canal_Temp[k, 8] / periodo_Canal_Temp_Unid[k, 8]);
                mod_9 = (periodo_Canal_Temp_Unid[k, 9] == 0 ? 0 : periodo_Canal_Temp[k, 9] / periodo_Canal_Temp_Unid[k, 9]);
                mod_10 = (periodo_Canal_Temp_Unid[k, 10] == 0 ? 0 : periodo_Canal_Temp[k, 10] / periodo_Canal_Temp_Unid[k, 10]);
                mod_11 = (periodo_Canal_Temp_Unid[k, 11] == 0 ? 0 : periodo_Canal_Temp[k, 11] / periodo_Canal_Temp_Unid[k, 11]);
                mod_12 = (periodo_Canal_Temp_Unid[k, 12] == 0 ? 0 : periodo_Canal_Temp[k, 12] / periodo_Canal_Temp_Unid[k, 12]);
                mod_13 = (periodo_Canal_Temp_Unid[k, 13] == 0 ? 0 : periodo_Canal_Temp[k, 13] / periodo_Canal_Temp_Unid[k, 13]);

                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);
                Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, "0. Cosmeticos", Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);

                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", periodo_Canal_Temp[k, 1] / 12, periodo_Canal_Temp[k, 2] / 12, periodo_Canal_Temp[k, 3] / 12, periodo_Canal_Temp[k, 4] / 12, periodo_Canal_Temp[k, 5] / 12, periodo_Canal_Temp[k, 6] / 6, periodo_Canal_Temp[k, 7] / 6, periodo_Canal_Temp[k, 8] / 3, periodo_Canal_Temp[k, 9] / 3, periodo_Canal_Temp[k, 10] / 1, periodo_Canal_Temp[k, 11] / 1, periodo_Canal_Temp[k, 12] / NumMeses, periodo_Canal_Temp[k, 13] / NumMeses);

                // CALCULAR DOLARES SHARE (%)"

                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Valor.Clone();
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Valor.Clone();
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Hogar.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Valor.Clone();
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Hogar.Clone();
                }

                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;

                if (xMoneda == 1)
                {                    
                    t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Canal_Temp[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                    t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Canal_Temp[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                    t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Canal_Temp[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                    t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Canal_Temp[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                    t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Canal_Temp[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                    t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Canal_Temp[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                    t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Canal_Temp[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                    t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Canal_Temp[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                    t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Canal_Temp[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                    t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Canal_Temp[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                    t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Canal_Temp[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                    t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Canal_Temp[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                    t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Canal_Temp[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                }

                // CALCULAR VALORES PROMEDIO (HOG.)"                

                for (int i = 0; i < periodo_Modalidad_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoModalidad == periodo_Modalidad_Temporal[i, 0].ToString())
                    {                        
                        t1 = periodo_Modalidad_Temporal[i, 1] != 0 ? periodo_Canal_Temp[k, 1] / periodo_Modalidad_Temporal[i, 1] : 0;
                        t2 = periodo_Modalidad_Temporal[i, 2] != 0 ? periodo_Canal_Temp[k, 2] / periodo_Modalidad_Temporal[i, 2] : 0;
                        t3 = periodo_Modalidad_Temporal[i, 3] != 0 ? periodo_Canal_Temp[k, 3] / periodo_Modalidad_Temporal[i, 3] : 0;
                        t4 = periodo_Modalidad_Temporal[i, 4] != 0 ? periodo_Canal_Temp[k, 4] / periodo_Modalidad_Temporal[i, 4] : 0;
                        t5 = periodo_Modalidad_Temporal[i, 5] != 0 ? periodo_Canal_Temp[k, 5] / periodo_Modalidad_Temporal[i, 5] : 0;
                        t6 = periodo_Modalidad_Temporal[i, 6] != 0 ? periodo_Canal_Temp[k, 6] / periodo_Modalidad_Temporal[i, 6] : 0;
                        t7 = periodo_Modalidad_Temporal[i, 7] != 0 ? periodo_Canal_Temp[k, 7] / periodo_Modalidad_Temporal[i, 7] : 0;
                        t8 = periodo_Modalidad_Temporal[i, 8] != 0 ? periodo_Canal_Temp[k, 8] / periodo_Modalidad_Temporal[i, 8] : 0;
                        t9 = periodo_Modalidad_Temporal[i, 9] != 0 ? periodo_Canal_Temp[k, 9] / periodo_Modalidad_Temporal[i, 9] : 0;
                        t10 = periodo_Modalidad_Temporal[i, 10] != 0 ? periodo_Canal_Temp[k, 10] / periodo_Modalidad_Temporal[i, 10] : 0;
                        t11 = periodo_Modalidad_Temporal[i, 11] != 0 ? periodo_Canal_Temp[k, 11] / periodo_Modalidad_Temporal[i, 11] : 0;
                        t12 = periodo_Modalidad_Temporal[i, 12] != 0 ? periodo_Canal_Temp[k, 12] / periodo_Modalidad_Temporal[i, 12] : 0;
                        t13 = periodo_Modalidad_Temporal[i, 13] != 0 ? periodo_Canal_Temp[k, 13] / periodo_Modalidad_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, "0. Cosmeticos", VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado, VariableGM, "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIA Y NSE % SHARE
            if (V1 == "Consolidado")
            {
                periodo_Cat_Temp = (double[,])periodo_Total_Mercado_Categoria_Valor.Clone();
                periodo_Cate_NSE_Temp = (double[,])periodo_Total_Mercado_CateNSE_Hogar.Clone();
            }
            else if (V1 == "Lima")
            {
                periodo_Cat_Temp = (double[,])periodo_Total_Lima_Categoria_Valor.Clone();
                periodo_Cate_NSE_Temp = (double[,])periodo_Total_Lima_CateNSE_Hogar.Clone();
            }
            else
            {
                periodo_Cat_Temp = (double[,])periodo_Total_Ciudades_Categoria_Valor.Clone();
                periodo_Cate_NSE_Temp = (double[,])periodo_Total_Ciudades_CateNSE_Hogar.Clone();
            }

            // CODIGO RECORRIDO POR CATEGORIAS Y NSE                              
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_CATEGORIA_NSE"))
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
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_NSE_CAT_VAL[rows, i] = valor_1;
                        }
                        rows++;

                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;

                        CodigoNSE = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_NSE.GetLength(0); i++) //  
                        {
                            if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                            {
                                //Mercado = Codigo_Nombres_NSE[i, 1].Substring(3, Codigo_Nombres_NSE[i, 1].Length - 3); ;
                                V1_ = Codigo_Nombres_NSE[i, 1];
                                break;
                            }
                        }
                        CodigoCategoria = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Categoria.GetLength(0); i++) // 
                        {
                            if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                            {
                                //V1__ = Codigo_Nombres_Categoria[i, 1].Substring(4, Codigo_Nombres_Categoria[i, 1].Length - 4);
                                Mercado_ = Codigo_Nombres_Categoria[i, 1];
                                break;
                            }
                        }

                        // CODIGO RECORRIDO POR CATEGORIA Y NSE % SHARE
                        if (xMoneda == 1)
                        {                            
                            for (int k = 0; k < periodo_Cat_Temp.GetLength(0); k++) //FILAS
                            {
                                if (CodigoCategoria == periodo_Cat_Temp[k, 0].ToString())
                                {
                                    t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Cat_Temp[k, 1] * 100 : 0;
                                    t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Cat_Temp[k, 2] * 100 : 0;
                                    t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Cat_Temp[k, 3] * 100 : 0;
                                    t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Cat_Temp[k, 4] * 100 : 0;
                                    t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Cat_Temp[k, 5] * 100 : 0;
                                    t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Cat_Temp[k, 6] * 100 : 0;
                                    t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Cat_Temp[k, 7] * 100 : 0;
                                    t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Cat_Temp[k, 8] * 100 : 0;
                                    t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Cat_Temp[k, 9] * 100 : 0;
                                    t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Cat_Temp[k, 10] * 100 : 0;
                                    t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Cat_Temp[k, 11] * 100;
                                    t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Cat_Temp[k, 12] * 100 : 0;
                                    t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Cat_Temp[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, Mercado_, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }                            
                        }

                        // CALCULAR VALORES PROMEDIO (HOG.) MODALIDAD POR TIPOS" 
                        for (int i = 0; i < periodo_Cate_NSE_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoNSE == periodo_Cate_NSE_Temp[i, 0].ToString() && CodigoCategoria == periodo_Cate_NSE_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Cate_NSE_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Cate_NSE_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Cate_NSE_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Cate_NSE_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Cate_NSE_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Cate_NSE_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Cate_NSE_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Cate_NSE_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Cate_NSE_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Cate_NSE_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Cate_NSE_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Cate_NSE_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Cate_NSE_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Cate_NSE_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Cate_NSE_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Cate_NSE_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Cate_NSE_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Cate_NSE_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Cate_NSE_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Cate_NSE_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Cate_NSE_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Cate_NSE_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Cate_NSE_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Cate_NSE_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Cate_NSE_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Cate_NSE_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado_, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }
                    }                                  
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_CATEGORIA_NSE"))
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
                int rows = 0; int columnas = 0;
                double mod_1 = 0, mod_2 = 0, mod_3 = 0, mod_4 = 0, mod_5 = 0, mod_6 = 0, mod_7 = 0, mod_8 = 0, mod_9 = 0, mod_10 = 0 , mod_11 = 0, mod_12 = 0, mod_13 = 0;
                int CodNSE = 0, CodCat = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {                    
                    int cols = reader_1.FieldCount;
                    columnas = cols;                    
                    while (reader_1.Read())
                    {
                        CodNSE = int.Parse(reader_1[0].ToString());
                        for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                        {
                            if (CodNSE == int.Parse(Codigo_Nombres_NSE[i, 0]))
                            {
                                V1_ = Codigo_Nombres_NSE[i, 1];
                                break;
                            }
                        }
                        CodCat = int.Parse(reader_1[1].ToString());
                        for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                        {
                            if (CodCat == int.Parse(Codigo_Nombres_Categoria[i, 0]))
                            {
                                Mercado = Codigo_Nombres_Categoria[i, 1];
                                break;
                            }
                        }

                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = periodo_NSE_CAT_VAL[rows, i] / double.Parse(reader_1[i].ToString());
                            }
                            switch (i)
                            {
                                case 2: mod_1 = valor_1; break;
                                case 3: mod_2 = valor_1; break;
                                case 4: mod_3 = valor_1; break;
                                case 5: mod_4 = valor_1; break;
                                case 6: mod_5 = valor_1; break;
                                case 7: mod_6 = valor_1; break;
                                case 8: mod_7 = valor_1; break;
                                case 9: mod_8 = valor_1; break;
                                case 10: mod_9 = valor_1; break;
                                case 11: mod_10 = valor_1; break;
                                case 12: mod_11 = valor_1; break;
                                case 13: mod_12 = valor_1; break;
                                case 14: mod_13 = valor_1; break;
                            }
                        }                         
                        Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);
                        rows++;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPO Y NSE SHARE
            if (V1 == "Consolidado")
            {
                periodo_Tipo_Temp = (double[,])periodo_Total_Mercado_Tipo_Valor.Clone();
                periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Mercado_TipoNSE_Hogar.Clone();
            }
            else if (V1 == "Lima")
            {
                periodo_Tipo_Temp = (double[,])periodo_Total_Lima_Tipo_Valor.Clone();
                periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Lima_TipoNSE_Hogar.Clone();
            }
            else
            {
                periodo_Tipo_Temp = (double[,])periodo_Total_Ciudades_Tipo_Valor.Clone();
                periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Ciudades_TipoNSE_Hogar.Clone();
            }
            // CODIGO RECORRIDO POR TIPOS Y NSE - PPU
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_TIPO_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_NSE_TIPO_VAL[rows, i] = valor_1;
                        }
                        rows++;

                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;

                        CodigoNSE = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_NSE.GetLength(0); i++) //  
                        {
                            if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                            {
                                V1_ = Codigo_Nombres_NSE[i, 1];
                                break;
                            }
                        }
                        CodigoTipo = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // 
                        {
                            if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                            {
                                Mercado_ = Codigo_Nombres_Tipo[i, 1];
                                break;
                            }
                        }

                        if (xMoneda == 1)
                        {
                            for (int k = 0; k < periodo_Tipo_Temp.GetLength(0); k++) //FILAS
                            {
                                if (CodigoTipo == periodo_Tipo_Temp[k, 0].ToString())
                                {
                                    t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Tipo_Temp[k, 1] * 100 : 0;
                                    t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Tipo_Temp[k, 2] * 100 : 0;
                                    t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Tipo_Temp[k, 3] * 100 : 0;
                                    t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Tipo_Temp[k, 4] * 100 : 0;
                                    t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Tipo_Temp[k, 5] * 100 : 0;
                                    t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Tipo_Temp[k, 6] * 100 : 0;
                                    t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Tipo_Temp[k, 7] * 100 : 0;
                                    t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Tipo_Temp[k, 8] * 100 : 0;
                                    t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Tipo_Temp[k, 9] * 100 : 0;
                                    t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Tipo_Temp[k, 10] * 100 : 0;
                                    t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Tipo_Temp[k, 11] * 100;
                                    t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Tipo_Temp[k, 12] * 100 : 0;
                                    t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Tipo_Temp[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, Mercado_, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }                            
                        }

                        // CALCULAR VALORES PROMEDIO (HOG.) MODALIDAD POR TIPOS" 
                        for (int i = 0; i < periodo_Tipo_NSE_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoNSE == periodo_Tipo_NSE_Temp[i, 0].ToString() && CodigoTipo == periodo_Tipo_NSE_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Tipo_NSE_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Tipo_NSE_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Tipo_NSE_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Tipo_NSE_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Tipo_NSE_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Tipo_NSE_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Tipo_NSE_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Tipo_NSE_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Tipo_NSE_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Tipo_NSE_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Tipo_NSE_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Tipo_NSE_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Tipo_NSE_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Tipo_NSE_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Tipo_NSE_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Tipo_NSE_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Tipo_NSE_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Tipo_NSE_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Tipo_NSE_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Tipo_NSE_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Tipo_NSE_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Tipo_NSE_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Tipo_NSE_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Tipo_NSE_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Tipo_NSE_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Tipo_NSE_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado_, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }
                    }                    
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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
                int rows = 0; int columnas = 0;
                double mod_1 = 0, mod_2 = 0, mod_3 = 0, mod_4 = 0, mod_5 = 0, mod_6 = 0, mod_7 = 0, mod_8 = 0, mod_9 = 0, mod_10 = 0, mod_11 = 0, mod_12 = 0, mod_13 = 0;
                int CodNSE = 0, CodTip = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    //int xCodNSE = 0, xCodTIPO = 0;
                    //int xFilaNSE = 0, xFilaTIPO = 0;
                    while (reader_1.Read())
                    {
                        CodNSE = int.Parse(reader_1[0].ToString());
                        CodTip = int.Parse(reader_1[1].ToString());       

                        for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE / 2 PQ TIENE 2 DIMENSIONES
                        {
                            if (CodNSE == int.Parse(Codigo_Nombres_NSE[i, 0]))
                            {
                                V1_ = Codigo_Nombres_NSE[i, 1];
                                break;
                            }     
                        }
                                               
                        for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE / 2 PQ TIENE 2 DIMENSIONES
                        {
                            if (CodTip == int.Parse(Codigo_Nombres_Tipo[i, 0]))
                            {
                                Mercado = Codigo_Nombres_Tipo[i, 1];
                                break;
                            }
                        }

                        for (int i = 2; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = periodo_NSE_TIPO_VAL[rows, i] / double.Parse(reader_1[i].ToString());
                            }
                            switch (i)
                            {
                                case 2: mod_1 = valor_1; break;
                                case 3: mod_2 = valor_1; break;
                                case 4: mod_3 = valor_1; break;
                                case 5: mod_4 = valor_1; break;
                                case 6: mod_5 = valor_1; break;
                                case 7: mod_6 = valor_1; break;
                                case 8: mod_7 = valor_1; break;
                                case 9: mod_8 = valor_1; break;
                                case 10: mod_9 = valor_1; break;
                                case 11: mod_10 = valor_1; break;
                                case 12: mod_11 = valor_1; break;
                                case 13: mod_12 = valor_1; break;
                                case 14: mod_13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);
                        rows++;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPO Y MODALIDAD SHARE
            if (V1 == "Consolidado")
            {
                periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Valor.Clone();
                periodo_Tipo_Temp = (double[,])periodo_Total_Mercado_Tipo_Valor.Clone();
                periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Mercado_TipoMOD_Hogar.Clone();
            }
            else if (V1 == "Lima")
            {
                periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Valor.Clone();
                periodo_Tipo_Temp = (double[,])periodo_Total_Lima_Tipo_Valor.Clone();
                periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Lima_TipoMOD_Hogar.Clone();
            }
            else
            {
                periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Valor.Clone();
                periodo_Tipo_Temp = (double[,])periodo_Total_Ciudades_Tipo_Valor.Clone();
                periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Ciudades_TipoMOD_Hogar.Clone();
            }
            // CODIGO RECORRIDO POR TIPOS Y MODALIDAD - PPU
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_TIPO_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value) {
                                valor_1 = 0;
                            }
                            else {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_TIPO_MOD_VAL[rows, i] = valor_1;
                        }
                        rows++;

                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        CodigoModalidad = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_Modalidad.GetLength(0); i++) //  
                        {
                            if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                            {
                                V1__ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                                Mercado = Codigo_Nombres_Modalidad[i, 1];
                                break;
                            }
                        }
                        CodigoTipo = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // 
                        {
                            if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                            {
                                V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                                Mercado_ = Codigo_Nombres_Tipo[i, 1];
                                break;
                            }
                        }

                        // CODIGO RECORRIDO POR TIPO Y MODALIDAD % SHARE
                        if (xMoneda == 1)
                        {                            
                            // RECORRIDO MATRIZ MODALIDAD
                            for (int k = 0; k < periodo_Modalidad_Temporal.GetLength(0); k++) //FILAS
                            {
                                if (CodigoModalidad == periodo_Modalidad_Temporal[k, 0].ToString())
                                {
                                    t1 = double.Parse(reader_1[2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / periodo_Modalidad_Temporal[k, 1] * 100 : 0;
                                    t2 = double.Parse(reader_1[3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / periodo_Modalidad_Temporal[k, 2] * 100 : 0;
                                    t3 = double.Parse(reader_1[4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / periodo_Modalidad_Temporal[k, 3] * 100 : 0;
                                    t4 = double.Parse(reader_1[5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / periodo_Modalidad_Temporal[k, 4] * 100 : 0;
                                    t5 = double.Parse(reader_1[6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / periodo_Modalidad_Temporal[k, 5] * 100 : 0;
                                    t6 = double.Parse(reader_1[7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / periodo_Modalidad_Temporal[k, 6] * 100 : 0;
                                    t7 = double.Parse(reader_1[8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / periodo_Modalidad_Temporal[k, 7] * 100 : 0;
                                    t8 = double.Parse(reader_1[9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / periodo_Modalidad_Temporal[k, 8] * 100 : 0;
                                    t9 = double.Parse(reader_1[10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / periodo_Modalidad_Temporal[k, 9] * 100 : 0;
                                    t10 = double.Parse(reader_1[11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / periodo_Modalidad_Temporal[k, 10] * 100 : 0;
                                    t11 = double.Parse(reader_1[12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / periodo_Modalidad_Temporal[k, 11] * 100 : 0;
                                    t12 = double.Parse(reader_1[13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / periodo_Modalidad_Temporal[k, 12] * 100 : 0;
                                    t13 = double.Parse(reader_1[14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / periodo_Modalidad_Temporal[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }

                            // RECORRIDO MATRIZ TIPOS % SHARE
                            for (int k = 0; k < periodo_Tipo_Temp.GetLength(0); k++) //FILAS
                            {
                                if (CodigoTipo == periodo_Tipo_Temp[k, 0].ToString())
                                {

                                    t1 = double.Parse(reader_1[2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / periodo_Tipo_Temp[k, 1] * 100 : 0;
                                    t2 = double.Parse(reader_1[3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / periodo_Tipo_Temp[k, 2] * 100 : 0;
                                    t3 = double.Parse(reader_1[4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / periodo_Tipo_Temp[k, 3] * 100 : 0;
                                    t4 = double.Parse(reader_1[5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / periodo_Tipo_Temp[k, 4] * 100 : 0;
                                    t5 = double.Parse(reader_1[6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / periodo_Tipo_Temp[k, 5] * 100 : 0;
                                    t6 = double.Parse(reader_1[7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / periodo_Tipo_Temp[k, 6] * 100 : 0;
                                    t7 = double.Parse(reader_1[8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / periodo_Tipo_Temp[k, 7] * 100 : 0;
                                    t8 = double.Parse(reader_1[9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / periodo_Tipo_Temp[k, 8] * 100 : 0;
                                    t9 = double.Parse(reader_1[10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / periodo_Tipo_Temp[k, 9] * 100 : 0;
                                    t10 = double.Parse(reader_1[11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / periodo_Tipo_Temp[k, 10] * 100 : 0;
                                    t11 = double.Parse(reader_1[12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / periodo_Tipo_Temp[k, 11] * 100 : 0;
                                    t12 = double.Parse(reader_1[13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / periodo_Tipo_Temp[k, 12] * 100 : 0;
                                    t13 = double.Parse(reader_1[14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / periodo_Tipo_Temp[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1__, "%", "SHARE DOLARES", Ciudad_, Mercado_, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }                            
                        }

                        // CALCULAR VALORES PROMEDIO (HOG.) MODALIDAD POR TIPOS"                        
                        for (int i = 0; i < periodo_Tipo_MOD_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoModalidad == periodo_Tipo_MOD_Temp[i, 0].ToString() && CodigoTipo == periodo_Tipo_MOD_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Tipo_MOD_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Tipo_MOD_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Tipo_MOD_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Tipo_MOD_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Tipo_MOD_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Tipo_MOD_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Tipo_MOD_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Tipo_MOD_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Tipo_MOD_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Tipo_MOD_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Tipo_MOD_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Tipo_MOD_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Tipo_MOD_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Tipo_MOD_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Tipo_MOD_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Tipo_MOD_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Tipo_MOD_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Tipo_MOD_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Tipo_MOD_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Tipo_MOD_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Tipo_MOD_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Tipo_MOD_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Tipo_MOD_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Tipo_MOD_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Tipo_MOD_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Tipo_MOD_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1__, "Suma", VariableGM, Ciudad_, Mercado_, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                break;
                            }
                        }
                    }                                   
                }                               
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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
                int rows = 0; int columnas = 0;
                double mod_1 = 0, mod_2 = 0, mod_3 = 0, mod_4 = 0, mod_5 = 0, mod_6 = 0, mod_7 = 0, mod_8 = 0, mod_9 = 0, mod_10 = 0, mod_11 = 0, mod_12 = 0, mod_13 = 0;

                int CodMod = 0, CodTip = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;

                    while (reader_1.Read())
                    {
                        CodMod = int.Parse(reader_1[0].ToString());
                        CodTip = int.Parse(reader_1[1].ToString());

                        for (int i = 0; i < Codigo_Nombres_Modalidad.GetLength(0); i++) // 
                        {
                            if (CodMod == int.Parse(Codigo_Nombres_Modalidad[i, 0]))
                            {                               
                                V1__ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3);
                                Mercado_ = Codigo_Nombres_Modalidad[i, 1];
                                break;
                            }
                        }

                        for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // 
                        {
                            if (CodTip == int.Parse(Codigo_Nombres_Tipo[i, 0]))
                            {
                                Mercado = Codigo_Nombres_Tipo[i, 1];    
                                V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                                break;
                            }
                        }

                        for (int i = 1; i < cols; i++) // ALMACENANDO DATOS UNIDADES EN ARRAY
                        {
                            if (reader_1[i] == DBNull.Value) {
                                valor_1 = 0; }
                            else {
                                valor_1 = double.Parse(reader_1[i].ToString()); }

                            if (V1 == "Consolidado") {
                                _Tipo_Canal_Mercado_Unidad[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima") {
                                _Tipo_Canal_Lima_Unidad[rows, i - 1] = valor_1;                                
                            }
                            else {
                                _Tipo_Canal_Ciudad_Unidad[rows, i - 1] = valor_1;                                
                            }
                        }

                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value) {
                                valor_1 = 0;
                            }
                            else {
                                valor_1 = periodo_TIPO_MOD_VAL[rows, i] / double.Parse(reader_1[i].ToString());
                            }

                            switch (i)
                            {
                                case 2: mod_1 = valor_1; break;
                                case 3: mod_2 = valor_1; break;
                                case 4: mod_3 = valor_1; break;
                                case 5: mod_4 = valor_1; break;
                                case 6: mod_5 = valor_1; break;
                                case 7: mod_6 = valor_1; break;
                                case 8: mod_7 = valor_1; break;
                                case 9: mod_8 = valor_1; break;
                                case 10: mod_9 = valor_1; break;
                                case 11: mod_10 = valor_1; break;
                                case 12: mod_11 = valor_1; break;
                                case 13: mod_12 = valor_1; break;
                                case 14: mod_13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(V1__, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);
                        Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado_, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);

                        rows++;
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIA Y MODALIDAD % SHARE
            if (V1 == "Consolidado")
            {
                periodo_Canal_Temp = (double[,])periodo_Total_Mercado_Modalidad_Valor.Clone();
                periodo_Cat_Temp = (double[,])periodo_Total_Mercado_Categoria_Valor.Clone();
                periodo_Cate_MOD_Temp = (double[,])periodo_Total_Mercado_CateMOD_Hogar.Clone();
            }
            else if (V1 == "Lima")
            {
                periodo_Canal_Temp = (double[,])periodo_Total_Lima_Modalidad_Valor.Clone();
                periodo_Cat_Temp = (double[,])periodo_Total_Lima_Categoria_Valor.Clone();
                periodo_Cate_MOD_Temp = (double[,])periodo_Total_Lima_CateMOD_Hogar.Clone();
            }
            else
            {
                periodo_Canal_Temp = (double[,])periodo_Total_Ciudades_Modalidad_Valor.Clone();
                periodo_Cat_Temp = (double[,])periodo_Total_Ciudades_Categoria_Valor.Clone();
                periodo_Cate_MOD_Temp = (double[,])periodo_Total_Ciudades_CateMOD_Hogar.Clone();
            }
            // CODIGO RECORRIDO POR CATEGORIA Y MODALIDAD                             
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_CATEGORIA_MODALIDAD"))
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
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_CAT_MOD_VAL[rows, i] = valor_1;
                        }
                        rows++;

                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;

                        CodigoModalidad = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_Modalidad.GetLength(0); i++) //  
                        {
                            if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                            {
                                V1__ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                                Mercado = Codigo_Nombres_Modalidad[i, 1];
                                break;
                            }
                        }
                        CodigoCategoria = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Categoria.GetLength(0); i++) // 
                        {
                            if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                            {
                                V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3);
                                Mercado_ = Codigo_Nombres_Categoria[i, 1];
                                break;
                            }
                        }

                        if (xMoneda==1)
                        {                            
                            // % SHARE CATEGORIA POR MODALIDAD SOBRE MODALIDAD
                            for (int k = 0; k < periodo_Canal_Temp.GetLength(0); k++) //FILAS
                            {
                                if (CodigoModalidad == periodo_Canal_Temp[k, 0].ToString())
                                {
                                    t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Canal_Temp[k, 1] * 100 : 0;
                                    t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Canal_Temp[k, 2] * 100 : 0;
                                    t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Canal_Temp[k, 3] * 100 : 0;
                                    t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Canal_Temp[k, 4] * 100 : 0;
                                    t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Canal_Temp[k, 5] * 100 : 0;
                                    t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Canal_Temp[k, 6] * 100 : 0;
                                    t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Canal_Temp[k, 7] * 100 : 0;
                                    t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Canal_Temp[k, 8] * 100 : 0;
                                    t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Canal_Temp[k, 9] * 100 : 0;
                                    t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Canal_Temp[k, 10] * 100 : 0;
                                    t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Canal_Temp[k, 11] * 100;
                                    t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Canal_Temp[k, 12] * 100 : 0;
                                    t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Canal_Temp[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1_, "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }
                            // % SHARE CATEGORIA POR MODALIDAD SOBRE CATEGORIA
                            for (int k = 0; k < periodo_Cat_Temp.GetLength(0); k++) //FILAS
                            {
                                if (CodigoCategoria == periodo_Cat_Temp[k, 0].ToString())
                                {
                                    t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Cat_Temp[k, 1] * 100 : 0;
                                    t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Cat_Temp[k, 2] * 100 : 0;
                                    t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Cat_Temp[k, 3] * 100 : 0;
                                    t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Cat_Temp[k, 4] * 100 : 0;
                                    t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Cat_Temp[k, 5] * 100 : 0;
                                    t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Cat_Temp[k, 6] * 100 : 0;
                                    t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Cat_Temp[k, 7] * 100 : 0;
                                    t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Cat_Temp[k, 8] * 100 : 0;
                                    t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Cat_Temp[k, 9] * 100 : 0;
                                    t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Cat_Temp[k, 10] * 100 : 0;
                                    t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Cat_Temp[k, 11] * 100;
                                    t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Cat_Temp[k, 12] * 100 : 0;
                                    t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Cat_Temp[k, 13] * 100 : 0;

                                    Actualizar_BD_Penetracion(V1__, "%", "SHARE DOLARES", Ciudad_, Mercado_, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                    break;
                                }
                            }
                        }


                        // CALCULAR VALORES PROMEDIO (HOG.) MODALIDAD POR CATEGORIA" 
                        for (int i = 0; i < periodo_Cate_MOD_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoModalidad == periodo_Cate_MOD_Temp[i, 0].ToString() && CodigoCategoria == periodo_Cate_MOD_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Cate_MOD_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Cate_MOD_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Cate_MOD_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Cate_MOD_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Cate_MOD_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Cate_MOD_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Cate_MOD_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Cate_MOD_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Cate_MOD_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Cate_MOD_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Cate_MOD_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Cate_MOD_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Cate_MOD_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Cate_MOD_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Cate_MOD_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Cate_MOD_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Cate_MOD_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Cate_MOD_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Cate_MOD_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Cate_MOD_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Cate_MOD_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Cate_MOD_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Cate_MOD_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Cate_MOD_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Cate_MOD_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Cate_MOD_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1_, "Suma", VariableGM, Ciudad_, Mercado, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                Actualizar_BD(V1__, "Suma", VariableGM, Ciudad_, Mercado_, VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                break;
                            }
                        }

                    }
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_CATEGORIA_MODALIDAD"))
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
                int rows = 0; int columnas = 0;
                double mod_1 = 0, mod_2 = 0, mod_3 = 0, mod_4 = 0, mod_5 = 0, mod_6 = 0, mod_7 = 0, mod_8 = 0, mod_9 = 0, mod_10 = 0, mod_11 = 0, mod_12 = 0, mod_13 = 0;
                int CodMod = 0, CodCat = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        CodMod = int.Parse(reader_1[0].ToString());
                        CodCat = int.Parse(reader_1[1].ToString());
                        for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE / 2 PQ TIENE 2 DIMENSIONES
                        {
                            if (CodMod == int.Parse(Codigo_Nombres_Modalidad[i, 0]))
                            {
                                V1__ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3);
                                Mercado_ = Codigo_Nombres_Modalidad[i, 1];
                                break;
                            }
                        }

                        for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE / 2 PQ TIENE 2 DIMENS
                        {
                            if (CodCat == int.Parse(Codigo_Nombres_Categoria[i, 0]))
                            {
                                Mercado = Codigo_Nombres_Categoria[i, 1];
                                V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3);
                                break;
                            }
                        }

                        for (int i = 0; i < cols ; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = periodo_CAT_MOD_VAL[rows, i] / double.Parse(reader_1[i].ToString());
                            }
                            switch (i)
                            {
                                case 2: mod_1 = valor_1; break;
                                case 3: mod_2 = valor_1; break;
                                case 4: mod_3 = valor_1; break;
                                case 5: mod_4 = valor_1; break;
                                case 6: mod_5 = valor_1; break;
                                case 7: mod_6 = valor_1; break;
                                case 8: mod_7 = valor_1; break;
                                case 9: mod_8 = valor_1; break;
                                case 10: mod_9 = valor_1; break;
                                case 11: mod_10 = valor_1; break;
                                case 12: mod_11 = valor_1; break;
                                case 13: mod_12 = valor_1; break;
                                case 14: mod_13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(V1__, "Suma", Unidad, Ciudad_, Mercado, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);
                        Actualizar_BD(V1_, "Suma", Unidad, Ciudad_, Mercado_, Unidad, "MENSUAL", mod_1, mod_2, mod_3, mod_4, mod_5, mod_6, mod_7, mod_8, mod_9, mod_10, mod_11, mod_12, mod_13);

                        rows++;
                    }
                }
            }
        }
        
        private int[] CompararArray(int[] x1, int[,] x2) 
        {
            //Procedimiento para validar si dos arreglos tienen los mismos valores.
            Boolean isArrayEqual = true;
            isArrayEqual = x1.Equals(x2);
            int[] faltantes = new int[15];
            int row = 0;
            if (x1.Length != (x2.Length/2)/5)
            {
                for (int i = 0; i < x1.Length - 1; i++)
                {
                    for (int u = 0; u < (x2.Length / 2) - 1; u++)
                    {
                        faltantes[row] = x2[u, 1];
                    }
                }                
            }
            return faltantes;
        }

        // UNIDADES
        public void Periodos_Cosmeticos_Total_Unidades(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, [Optional] int NumMeses)
        {

            Variable = "UNIDADES";
            Variable_Promedio = "UNIDADES MES";
            V2 = "";

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

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_MERCADO"))
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
                                periodo_Total_Mercado_Unidad[rows, i - 1] = valor_1;
                                periodo_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Unidad[rows, i - 1] = valor_1;
                                periodo_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Unidad[rows, i - 1] = valor_1;
                                periodo_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }

                    Actualizar_BD(V1, "Suma", Variable, "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_Temp_Unidad[0, 0], periodo_Temp_Unidad[0, 1], periodo_Temp_Unidad[0, 2], periodo_Temp_Unidad[0, 3], periodo_Temp_Unidad[0, 4], periodo_Temp_Unidad[0, 5], periodo_Temp_Unidad[0, 6], periodo_Temp_Unidad[0, 7], periodo_Temp_Unidad[0, 8], periodo_Temp_Unidad[0, 9], periodo_Temp_Unidad[0, 10], periodo_Temp_Unidad[0, 11], periodo_Temp_Unidad[0, 12]);
                    // PROMEDIO MENSUAL
                    Actualizar_BD(V1, "Suma", Variable_Promedio, "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_Temp_Unidad[0, 0] / 12, periodo_Temp_Unidad[0, 1] / 12, periodo_Temp_Unidad[0, 2] / 12, periodo_Temp_Unidad[0, 3] / 12, periodo_Temp_Unidad[0, 4] / 12, periodo_Temp_Unidad[0, 5] / 6, periodo_Temp_Unidad[0, 6] / 6, periodo_Temp_Unidad[0, 7] / 3, periodo_Temp_Unidad[0, 8] / 3, periodo_Temp_Unidad[0, 9] / 1, periodo_Temp_Unidad[0, 10] / 1, periodo_Temp_Unidad[0, 11] / NumMeses, periodo_Temp_Unidad[0, 12] / NumMeses);

                    if (xCiudad != "1,2,5") //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS
                    {
                        Actualizar_BD("Cosmeticos", "Suma", Variable, Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_Temp_Unidad[0, 0], periodo_Temp_Unidad[0, 1], periodo_Temp_Unidad[0, 2], periodo_Temp_Unidad[0, 3], periodo_Temp_Unidad[0, 4], periodo_Temp_Unidad[0, 5], periodo_Temp_Unidad[0, 6], periodo_Temp_Unidad[0, 7], periodo_Temp_Unidad[0, 8], periodo_Temp_Unidad[0, 9], periodo_Temp_Unidad[0, 10], periodo_Temp_Unidad[0, 11], periodo_Temp_Unidad[0, 12]);
                        // PROMEDIO MENSUAL
                        Actualizar_BD("Cosmeticos", "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_Temp_Unidad[0, 0] / 12, periodo_Temp_Unidad[0, 1] / 12, periodo_Temp_Unidad[0, 2] / 12, periodo_Temp_Unidad[0, 3] / 12, periodo_Temp_Unidad[0, 4] / 12, periodo_Temp_Unidad[0, 5] / 6, periodo_Temp_Unidad[0, 6] / 6, periodo_Temp_Unidad[0, 7] / 3, periodo_Temp_Unidad[0, 8] / 3, periodo_Temp_Unidad[0, 9] / 1, periodo_Temp_Unidad[0, 10] / 1, periodo_Temp_Unidad[0, 11] / NumMeses, periodo_Temp_Unidad[0, 12] / NumMeses);
                    }

                    // CALCULAR UNIDADES PROMEDIO (HOG.)"
                    if (V1 == "Consolidado")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Hogar.Clone();
                    }
                    else if (V1 == "Lima")
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Lima_Hogar.Clone();
                    }
                    else
                    {
                        periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Hogar.Clone();
                    }

                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = double.Parse(periodo_Total_Temporal[0, 0].ToString()) != 0 ? periodo_Temp_Unidad[0, 0] / periodo_Total_Temporal[0, 0] : 0;
                    t2 = double.Parse(periodo_Total_Temporal[0, 1].ToString()) != 0 ? periodo_Temp_Unidad[0, 1] / periodo_Total_Temporal[0, 1] : 0;
                    t3 = double.Parse(periodo_Total_Temporal[0, 2].ToString()) != 0 ? periodo_Temp_Unidad[0, 2] / periodo_Total_Temporal[0, 2] : 0;
                    t4 = double.Parse(periodo_Total_Temporal[0, 3].ToString()) != 0 ? periodo_Temp_Unidad[0, 3] / periodo_Total_Temporal[0, 3] : 0;
                    t5 = double.Parse(periodo_Total_Temporal[0, 4].ToString()) != 0 ? periodo_Temp_Unidad[0, 4] / periodo_Total_Temporal[0, 4] : 0;
                    t6 = double.Parse(periodo_Total_Temporal[0, 5].ToString()) != 0 ? periodo_Temp_Unidad[0, 5] / periodo_Total_Temporal[0, 5] : 0;
                    t7 = double.Parse(periodo_Total_Temporal[0, 6].ToString()) != 0 ? periodo_Temp_Unidad[0, 6] / periodo_Total_Temporal[0, 6] : 0;
                    t8 = double.Parse(periodo_Total_Temporal[0, 7].ToString()) != 0 ? periodo_Temp_Unidad[0, 7] / periodo_Total_Temporal[0, 7] : 0;
                    t9 = double.Parse(periodo_Total_Temporal[0, 8].ToString()) != 0 ? periodo_Temp_Unidad[0, 8] / periodo_Total_Temporal[0, 8] : 0;
                    t10 = double.Parse(periodo_Total_Temporal[0, 9].ToString()) != 0 ? periodo_Temp_Unidad[0, 9] / periodo_Total_Temporal[0, 9] : 0;
                    t11 = double.Parse(periodo_Total_Temporal[0, 10].ToString()) != 0 ? periodo_Temp_Unidad[0, 10] / periodo_Total_Temporal[0, 10] : 0;
                    t12 = double.Parse(periodo_Total_Temporal[0, 11].ToString()) != 0 ? periodo_Temp_Unidad[0, 11] / periodo_Total_Temporal[0, 11] : 0;
                    t13 = double.Parse(periodo_Total_Temporal[0, 12].ToString()) != 0 ? periodo_Temp_Unidad[0, 12] / periodo_Total_Temporal[0, 12] : 0;

                    Actualizar_BD(V1, "Suma", "UNIDADES PROMEDIO (HOG.)", "0. Consolidado", "0. Cosmeticos", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                    if (xCiudad != "1,2,5") //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS
                    { 
                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, "0. Cosmeticos", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        // CALCULAR UNIDADES SHARE (%)"
                        t1 = periodo_Temp_Unidad[0, 0] / periodo_Total_Mercado_Unidad[0, 0] * 100;
                        t2 = periodo_Temp_Unidad[0, 1] / periodo_Total_Mercado_Unidad[0, 1] * 100;
                        t3 = periodo_Temp_Unidad[0, 2] / periodo_Total_Mercado_Unidad[0, 2] * 100;
                        t4 = periodo_Temp_Unidad[0, 3] / periodo_Total_Mercado_Unidad[0, 3] * 100;
                        t5 = periodo_Temp_Unidad[0, 4] / periodo_Total_Mercado_Unidad[0, 4] * 100;
                        t6 = periodo_Temp_Unidad[0, 5] / periodo_Total_Mercado_Unidad[0, 5] * 100;
                        t7 = periodo_Temp_Unidad[0, 6] / periodo_Total_Mercado_Unidad[0, 6] * 100;
                        t8 = periodo_Temp_Unidad[0, 7] / periodo_Total_Mercado_Unidad[0, 7] * 100;
                        t9 = periodo_Temp_Unidad[0, 8] / periodo_Total_Mercado_Unidad[0, 8] * 100;
                        t10 = periodo_Temp_Unidad[0, 9] / periodo_Total_Mercado_Unidad[0, 9] * 100;
                        t11 = periodo_Temp_Unidad[0, 10] / periodo_Total_Mercado_Unidad[0, 10] * 100;
                        t12 = periodo_Temp_Unidad[0, 11] / periodo_Total_Mercado_Unidad[0, 11] * 100;
                        t13 = periodo_Temp_Unidad[0, 12] / periodo_Total_Mercado_Unidad[0, 12] * 100;

                        Actualizar_BD_Penetracion(V1, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2 , t3 , t4 , t5 , t6 , t7 , t8 , t9 , t10 , t11 , t12 , t13 );
                    }
                }
            }

            // CODIGO RECORRIDO POR NSE              
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_NSE"))
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
                                periodo_Total_Mercado_NSE_Unidad[rows, i - 1] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_NSE_Unidad[rows, i - 1] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_NSE_Unidad[rows, i - 1] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }

                    if (rows == 4)
                    {
                        periodo_NSE_Temp_Unidad[rows, 0] = 489; periodo_NSE_Temp_Unidad[rows, 1] = 0;
                        periodo_NSE_Temp_Unidad[rows, 2] = 0; periodo_NSE_Temp_Unidad[rows, 3] = 0;
                        periodo_NSE_Temp_Unidad[rows, 4] = 0; periodo_NSE_Temp_Unidad[rows, 5] = 0;
                        periodo_NSE_Temp_Unidad[rows, 6] = 0; periodo_NSE_Temp_Unidad[rows, 7] = 0;
                        periodo_NSE_Temp_Unidad[rows, 8] = 0; periodo_NSE_Temp_Unidad[rows, 9] = 0;
                        periodo_NSE_Temp_Unidad[rows, 10] = 0; periodo_NSE_Temp_Unidad[rows, 11] = 0;
                        periodo_NSE_Temp_Unidad[rows, 12] = 0; periodo_NSE_Temp_Unidad[rows, 13] = 0;
                    }
                }
            }
            string CodigoNSE;
            for (int k = 0; k < periodo_NSE_Temp_Unidad.GetLength(0); k++) //FILAS
            {
                CodigoNSE = periodo_NSE_Temp_Unidad[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        Mercado = Codigo_Nombres_NSE[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_NSE_Temp_Unidad[k, 1], periodo_NSE_Temp_Unidad[k, 2], periodo_NSE_Temp_Unidad[k, 3], periodo_NSE_Temp_Unidad[k, 4], periodo_NSE_Temp_Unidad[k, 5], periodo_NSE_Temp_Unidad[k, 6], periodo_NSE_Temp_Unidad[k, 7], periodo_NSE_Temp_Unidad[k, 8], periodo_NSE_Temp_Unidad[k, 9], periodo_NSE_Temp_Unidad[k, 10], periodo_NSE_Temp_Unidad[k, 11], periodo_NSE_Temp_Unidad[k, 12], periodo_NSE_Temp_Unidad[k, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", periodo_NSE_Temp_Unidad[k, 1] / 12, periodo_NSE_Temp_Unidad[k, 2] / 12, periodo_NSE_Temp_Unidad[k, 3] / 12, periodo_NSE_Temp_Unidad[k, 4] / 12, periodo_NSE_Temp_Unidad[k, 5] / 12, periodo_NSE_Temp_Unidad[k, 6] / 6, periodo_NSE_Temp_Unidad[k, 7] / 6, periodo_NSE_Temp_Unidad[k, 8] / 3, periodo_NSE_Temp_Unidad[k, 9] / 3, periodo_NSE_Temp_Unidad[k, 10] / 1, periodo_NSE_Temp_Unidad[k, 11] / 1, periodo_NSE_Temp_Unidad[k, 12] / NumMeses, periodo_NSE_Temp_Unidad[k, 13] / NumMeses);

                // CALCULAR UNIDADES SHARE (%)"
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Unidad.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Unidad.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Unidad.Clone();
                }

                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                t1 = double.Parse(periodo_Total_Temporal[0, 0].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 1] / periodo_Total_Temporal[0, 0] * 100: 0;
                t2 = double.Parse(periodo_Total_Temporal[0, 1].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                t3 = double.Parse(periodo_Total_Temporal[0, 2].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                t4 = double.Parse(periodo_Total_Temporal[0, 3].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                t5 = double.Parse(periodo_Total_Temporal[0, 4].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                t6 = double.Parse(periodo_Total_Temporal[0, 5].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                t7 = double.Parse(periodo_Total_Temporal[0, 6].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                t8 = double.Parse(periodo_Total_Temporal[0, 7].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                t9 = double.Parse(periodo_Total_Temporal[0, 8].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                t10 = double.Parse(periodo_Total_Temporal[0, 9].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                t11 = double.Parse(periodo_Total_Temporal[0, 10].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                t12 = double.Parse(periodo_Total_Temporal[0, 11].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                t13 = double.Parse(periodo_Total_Temporal[0, 12].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                // CALCULAR UNIDADES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Mercado_NSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Lima_NSE_Hogar.Clone();
                }
                else
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Ciudades_NSE_Hogar.Clone();
                }
                for (int i = 0; i < periodo_NSE_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoNSE == periodo_NSE_Temporal[i, 0].ToString())
                    {
                        //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_NSE_Temporal[i, 1].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 1] / periodo_NSE_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_NSE_Temporal[i, 2].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 2] / periodo_NSE_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_NSE_Temporal[i, 3].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 3] / periodo_NSE_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_NSE_Temporal[i, 4].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 4] / periodo_NSE_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_NSE_Temporal[i, 5].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 5] / periodo_NSE_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_NSE_Temporal[i, 6].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 6] / periodo_NSE_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_NSE_Temporal[i, 7].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 7] / periodo_NSE_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_NSE_Temporal[i, 8].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 8] / periodo_NSE_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_NSE_Temporal[i, 9].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 9] / periodo_NSE_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_NSE_Temporal[i, 10].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 10] / periodo_NSE_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_NSE_Temporal[i, 11].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 11] / periodo_NSE_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_NSE_Temporal[i, 12].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 12] / periodo_NSE_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_NSE_Temporal[i, 13].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 13] / periodo_NSE_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, "0. Cosmeticos", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }                   
                }
            }

            // CODIGO RECORRIDO POR CATEGORIAS           
            string CodigoCategoria;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_CATEGORIA"))
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
                                periodo_Total_Mercado_Categoria_Unidad[rows, i - 1] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_Unidad[rows, i - 1] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_Unidad[rows, i - 1] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }           
            for (int k = 0; k < periodo_Cat_Temp_Unid.GetLength(0); k++) //FILAS
            {
                CodigoCategoria = periodo_Cat_Temp_Unid[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Categoria[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Cat_Temp_Unid[k, 1], periodo_Cat_Temp_Unid[k, 2], periodo_Cat_Temp_Unid[k, 3], periodo_Cat_Temp_Unid[k, 4], periodo_Cat_Temp_Unid[k, 5], periodo_Cat_Temp_Unid[k, 6], periodo_Cat_Temp_Unid[k, 7], periodo_Cat_Temp_Unid[k, 8], periodo_Cat_Temp_Unid[k, 9], periodo_Cat_Temp_Unid[k, 10], periodo_Cat_Temp_Unid[k, 11], periodo_Cat_Temp_Unid[k, 12], periodo_Cat_Temp_Unid[k, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Cat_Temp_Unid[k, 1] / 12, periodo_Cat_Temp_Unid[k, 2] / 12, periodo_Cat_Temp_Unid[k, 3] / 12, periodo_Cat_Temp_Unid[k, 4] / 12, periodo_Cat_Temp_Unid[k, 5] / 12, periodo_Cat_Temp_Unid[k, 6] / 6, periodo_Cat_Temp_Unid[k, 7] / 6, periodo_Cat_Temp_Unid[k, 8] / 3, periodo_Cat_Temp_Unid[k, 9] / 3, periodo_Cat_Temp_Unid[k, 10] / 1, periodo_Cat_Temp_Unid[k, 11] / 1, periodo_Cat_Temp_Unid[k, 12] / NumMeses, periodo_Cat_Temp_Unid[k, 13]/NumMeses);

                // CALCULAR UNIDADES SHARE (%)"
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Unidad.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Unidad.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Unidad.Clone();
                }

                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                t1 = double.Parse(periodo_Total_Temporal[0, 0].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                t2 = double.Parse(periodo_Total_Temporal[0, 1].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                t3 = double.Parse(periodo_Total_Temporal[0, 2].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                t4 = double.Parse(periodo_Total_Temporal[0, 3].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                t5 = double.Parse(periodo_Total_Temporal[0, 4].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                t6 = double.Parse(periodo_Total_Temporal[0, 5].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                t7 = double.Parse(periodo_Total_Temporal[0, 6].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                t8 = double.Parse(periodo_Total_Temporal[0, 7].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                t9 = double.Parse(periodo_Total_Temporal[0, 8].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                t10 = double.Parse(periodo_Total_Temporal[0, 9].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                t11 = double.Parse(periodo_Total_Temporal[0, 10].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                t12 = double.Parse(periodo_Total_Temporal[0, 11].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                t13 = double.Parse(periodo_Total_Temporal[0, 12].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                // CALCULAR UNIDADES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Mercado_Categoria_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Lima_Categoria_Hogar.Clone();
                }
                else
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Ciudades_Categoria_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Cat_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoCategoria == periodo_Cat_Temporal[i, 0].ToString())
                    {
                        //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_Cat_Temporal[i, 1].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 1] / periodo_Cat_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_Cat_Temporal[i, 2].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 2] / periodo_Cat_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_Cat_Temporal[i, 3].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 3] / periodo_Cat_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_Cat_Temporal[i, 4].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 4] / periodo_Cat_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_Cat_Temporal[i, 5].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 5] / periodo_Cat_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_Cat_Temporal[i, 6].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 6] / periodo_Cat_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_Cat_Temporal[i, 7].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 7] / periodo_Cat_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_Cat_Temporal[i, 8].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 8] / periodo_Cat_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_Cat_Temporal[i, 9].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 9] / periodo_Cat_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_Cat_Temporal[i, 10].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 10] / periodo_Cat_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_Cat_Temporal[i, 11].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 11] / periodo_Cat_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_Cat_Temporal[i, 12].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 12] / periodo_Cat_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_Cat_Temporal[i, 13].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 13] / periodo_Cat_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, "0. Cosmeticos2", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        
                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPOS      
            string CodigoTipo;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO"))
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
                                periodo_Total_Mercado_Tipo_Unidad[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Tipo_Unidad[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Tipo_Unidad[rows, i - 1] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Tipo_Temp_Unid.GetLength(0); ii++) // FILAS
            {
                CodigoTipo = periodo_Tipo_Temp_Unid[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                    {
                        V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado = Codigo_Nombres_Tipo[i, 1];
                        break;
                    }
                }

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Tipo_Temp_Unid[ii, 1], periodo_Tipo_Temp_Unid[ii, 2], periodo_Tipo_Temp_Unid[ii, 3], periodo_Tipo_Temp_Unid[ii, 4], periodo_Tipo_Temp_Unid[ii, 5], periodo_Tipo_Temp_Unid[ii, 6], periodo_Tipo_Temp_Unid[ii, 7], periodo_Tipo_Temp_Unid[ii, 8], periodo_Tipo_Temp_Unid[ii, 9], periodo_Tipo_Temp_Unid[ii, 10], periodo_Tipo_Temp_Unid[ii, 11], periodo_Tipo_Temp_Unid[ii, 12], periodo_Tipo_Temp_Unid[ii, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Tipo_Temp_Unid[ii, 1] / 12, periodo_Tipo_Temp_Unid[ii, 2] / 12, periodo_Tipo_Temp_Unid[ii, 3] / 12, periodo_Tipo_Temp_Unid[ii, 4] / 12, periodo_Tipo_Temp_Unid[ii, 5] / 12, periodo_Tipo_Temp_Unid[ii, 6] / 6, periodo_Tipo_Temp_Unid[ii, 7] / 6, periodo_Tipo_Temp_Unid[ii, 8] / 3, periodo_Tipo_Temp_Unid[ii, 9] / 3, periodo_Tipo_Temp_Unid[ii, 10] / 1, periodo_Tipo_Temp_Unid[ii, 11] / 1, periodo_Tipo_Temp_Unid[ii, 12] / NumMeses, periodo_Tipo_Temp_Unid[ii, 13]/NumMeses);

                // CALCULAR UNIDADES SHARE (%)"
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Unidad.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Unidad.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Unidad.Clone();
                }

                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Tipo_Temp_Unid[ii, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Tipo_Temp_Unid[ii, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Tipo_Temp_Unid[ii, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Tipo_Temp_Unid[ii, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Tipo_Temp_Unid[ii, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Tipo_Temp_Unid[ii, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Tipo_Temp_Unid[ii, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Tipo_Temp_Unid[ii, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Tipo_Temp_Unid[ii, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Tipo_Temp_Unid[ii, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Tipo_Temp_Unid[ii, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Tipo_Temp_Unid[ii, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Tipo_Temp_Unid[ii, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                // CALCULAR UNIDADES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Mercado_Tipo_Hogar_.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Lima_Tipo_Hogar_.Clone();
                }
                else
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Ciudades_Tipo_Hogar_.Clone();
                }

                for (int i = 0; i < periodo_Tipo_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoTipo == periodo_Tipo_Temporal[i, 0].ToString())
                    {
                        //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_Tipo_Temporal[i, 1].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 1] / periodo_Tipo_Temporal[i, 1] : 0;
                        t2 = double.Parse(periodo_Tipo_Temporal[i, 2].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 2] / periodo_Tipo_Temporal[i, 2] : 0;
                        t3 = double.Parse(periodo_Tipo_Temporal[i, 3].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 3] / periodo_Tipo_Temporal[i, 3] : 0;
                        t4 = double.Parse(periodo_Tipo_Temporal[i, 4].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 4] / periodo_Tipo_Temporal[i, 4] : 0;
                        t5 = double.Parse(periodo_Tipo_Temporal[i, 5].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 5] / periodo_Tipo_Temporal[i, 5] : 0;
                        t6 = double.Parse(periodo_Tipo_Temporal[i, 6].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 6] / periodo_Tipo_Temporal[i, 6] : 0;
                        t7 = double.Parse(periodo_Tipo_Temporal[i, 7].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 7] / periodo_Tipo_Temporal[i, 7] : 0;
                        t8 = double.Parse(periodo_Tipo_Temporal[i, 8].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 8] / periodo_Tipo_Temporal[i, 8] : 0;
                        t9 = double.Parse(periodo_Tipo_Temporal[i, 9].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 9] / periodo_Tipo_Temporal[i, 9] : 0;
                        t10 = double.Parse(periodo_Tipo_Temporal[i, 10].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 10] / periodo_Tipo_Temporal[i, 10] : 0;
                        t11 = double.Parse(periodo_Tipo_Temporal[i, 11].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 11] / periodo_Tipo_Temporal[i, 11] : 0;
                        t12 = double.Parse(periodo_Tipo_Temporal[i, 12].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 12] / periodo_Tipo_Temporal[i, 12] : 0;
                        t13 = double.Parse(periodo_Tipo_Temporal[i, 13].ToString()) != 0 ? periodo_Tipo_Temp_Unid[ii, 13] / periodo_Tipo_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, "0. Cosmeticos2", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR MODALIDAD      
            string CodigoModalidad;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MODALIDAD", DbType.String, Codigo_cadena_Modalidad);
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
                                periodo_Total_Mercado_Modalidad_Unidad[rows, i - 1] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Modalidad_Unidad[rows, i - 1] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Modalidad_Unidad[rows, i - 1] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int k = 0; k < periodo_Canal_Temp_Unid.GetLength(0); k++) //FILAS
            {
                CodigoModalidad = periodo_Canal_Temp_Unid[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                    {
                        V1_ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Modalidad[i, 1];
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Canal_Temp_Unid[k, 1], periodo_Canal_Temp_Unid[k, 2], periodo_Canal_Temp_Unid[k, 3], periodo_Canal_Temp_Unid[k, 4], periodo_Canal_Temp_Unid[k, 5], periodo_Canal_Temp_Unid[k, 6], periodo_Canal_Temp_Unid[k, 7], periodo_Canal_Temp_Unid[k, 8], periodo_Canal_Temp_Unid[k, 9], periodo_Canal_Temp_Unid[k, 10], periodo_Canal_Temp_Unid[k, 11], periodo_Canal_Temp_Unid[k, 12], periodo_Canal_Temp_Unid[k, 13]);
                // PROMEDIO MENSUAL
                Actualizar_BD(V1_, "Suma", Variable_Promedio, Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", periodo_Canal_Temp_Unid[k, 1] / 12, periodo_Canal_Temp_Unid[k, 2] / 12, periodo_Canal_Temp_Unid[k, 3] / 12, periodo_Canal_Temp_Unid[k, 4] / 12, periodo_Canal_Temp_Unid[k, 5] / 12, periodo_Canal_Temp_Unid[k, 6] / 6, periodo_Canal_Temp_Unid[k, 7] / 6, periodo_Canal_Temp_Unid[k, 8] / 3, periodo_Canal_Temp_Unid[k, 9] / 3, periodo_Canal_Temp_Unid[k, 10] / 1, periodo_Canal_Temp_Unid[k, 11] / 1, periodo_Canal_Temp_Unid[k, 12] / NumMeses, periodo_Canal_Temp_Unid[k, 13]/NumMeses);

                // CALCULAR UNIDADES SHARE (%)"
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Unidad.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Unidad.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Unidad.Clone();
                }

                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Canal_Temp_Unid[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Canal_Temp_Unid[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Canal_Temp_Unid[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Canal_Temp_Unid[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Canal_Temp_Unid[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Canal_Temp_Unid[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Canal_Temp_Unid[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Canal_Temp_Unid[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Canal_Temp_Unid[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Canal_Temp_Unid[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Canal_Temp_Unid[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Canal_Temp_Unid[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Canal_Temp_Unid[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                // CALCULAR UNIDADES PROMEDIO (HOG.)"
                if (V1 == "Consolidado")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Hogar.Clone();
                }
                else
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Modalidad_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoModalidad == periodo_Modalidad_Temporal[i, 0].ToString())
                    {
                        //double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = periodo_Modalidad_Temporal[i, 1] != 0 ? periodo_Canal_Temp_Unid[k, 1] / periodo_Modalidad_Temporal[i, 1] : 0;
                        t2 = periodo_Modalidad_Temporal[i, 2] != 0 ? periodo_Canal_Temp_Unid[k, 2] / periodo_Modalidad_Temporal[i, 2] : 0;
                        t3 = periodo_Modalidad_Temporal[i, 3] != 0 ? periodo_Canal_Temp_Unid[k, 3] / periodo_Modalidad_Temporal[i, 3] : 0;
                        t4 = periodo_Modalidad_Temporal[i, 4] != 0 ? periodo_Canal_Temp_Unid[k, 4] / periodo_Modalidad_Temporal[i, 4] : 0;
                        t5 = periodo_Modalidad_Temporal[i, 5] != 0 ? periodo_Canal_Temp_Unid[k, 5] / periodo_Modalidad_Temporal[i, 5] : 0;
                        t6 = periodo_Modalidad_Temporal[i, 6] != 0 ? periodo_Canal_Temp_Unid[k, 6] / periodo_Modalidad_Temporal[i, 6] : 0;
                        t7 = periodo_Modalidad_Temporal[i, 7] != 0 ? periodo_Canal_Temp_Unid[k, 7] / periodo_Modalidad_Temporal[i, 7] : 0;
                        t8 = periodo_Modalidad_Temporal[i, 8] != 0 ? periodo_Canal_Temp_Unid[k, 8] / periodo_Modalidad_Temporal[i, 8] : 0;
                        t9 = periodo_Modalidad_Temporal[i, 9] != 0 ? periodo_Canal_Temp_Unid[k, 9] / periodo_Modalidad_Temporal[i, 9] : 0;
                        t10 = periodo_Modalidad_Temporal[i, 10] != 0 ? periodo_Canal_Temp_Unid[k, 10] / periodo_Modalidad_Temporal[i, 10] : 0;
                        t11 = periodo_Modalidad_Temporal[i, 11] != 0 ? periodo_Canal_Temp_Unid[k, 11] / periodo_Modalidad_Temporal[i, 11] : 0;
                        t12 = periodo_Modalidad_Temporal[i, 12] != 0 ? periodo_Canal_Temp_Unid[k, 12] / periodo_Modalidad_Temporal[i, 12] : 0;
                        t13 = periodo_Modalidad_Temporal[i, 13] != 0 ? periodo_Canal_Temp_Unid[k, 13] / periodo_Modalidad_Temporal[i, 13] : 0;

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, "0. Cosmeticos", "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1 * 1000, t2 * 1000, t3 * 1000, t4 * 1000, t5 * 1000, t6 * 1000, t7 * 1000, t8 * 1000, t9 * 1000, t10 * 1000, t11 * 1000, t12 * 1000, t13 * 1000);

                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPO Y MODALIDAD SHARE
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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

                int columnas = 0;

                if (V1 == "Consolidado")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Unidad.Clone();
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Mercado_Tipo_Unidad.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Unidad.Clone();
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Lima_Tipo_Unidad.Clone();
                }
                else
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Unidad.Clone();
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Ciudades_Tipo_Unidad.Clone();
                }               

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    while (reader_1.Read())
                    {
                        CodigoModalidad = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_Modalidad.GetLength(0); i++) //  
                        {
                            if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                            {
                                V1_ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                                Mercado = Codigo_Nombres_Modalidad[i, 1];
                            }
                        }
                        CodigoTipo = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // 
                        {
                            if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                            {
                                V1__ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                                Mercado_ = Codigo_Nombres_Tipo[i, 1];
                                break;
                            }
                        }
                        // RECORRIDO MATRIZ MODALIDAD
                        for (int k = 0; k < periodo_Modalidad_Temporal.GetLength(0); k++) //FILAS
                        {
                            if (CodigoModalidad == periodo_Modalidad_Temporal[k, 0].ToString())
                            {
                                t1 = double.Parse(reader_1[2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / periodo_Modalidad_Temporal[k, 1] * 100 : 0;
                                t2 = double.Parse(reader_1[3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / periodo_Modalidad_Temporal[k, 2] * 100 : 0;
                                t3 = double.Parse(reader_1[4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / periodo_Modalidad_Temporal[k, 3] * 100 : 0;
                                t4 = double.Parse(reader_1[5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / periodo_Modalidad_Temporal[k, 4] * 100 : 0;
                                t5 = double.Parse(reader_1[6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / periodo_Modalidad_Temporal[k, 5] * 100 : 0;
                                t6 = double.Parse(reader_1[7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / periodo_Modalidad_Temporal[k, 6] * 100 : 0;
                                t7 = double.Parse(reader_1[8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / periodo_Modalidad_Temporal[k, 7] * 100 : 0;
                                t8 = double.Parse(reader_1[9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / periodo_Modalidad_Temporal[k, 8] * 100 : 0;
                                t9 = double.Parse(reader_1[10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / periodo_Modalidad_Temporal[k, 9] * 100 : 0;
                                t10 = double.Parse(reader_1[11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / periodo_Modalidad_Temporal[k, 10] * 100 : 0;
                                t11 = double.Parse(reader_1[12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / periodo_Modalidad_Temporal[k, 11] * 100 : 0;
                                t12 = double.Parse(reader_1[13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / periodo_Modalidad_Temporal[k, 12] * 100 : 0;
                                t13 = double.Parse(reader_1[14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / periodo_Modalidad_Temporal[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(V1__, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }

                        // RECORRIDO MATRIZ TIPOS
                        if (V1 == "Consolidado")
                        {
                            periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Mercado_TipoMOD_Hogar.Clone();
                        }
                        else if (V1 == "Lima")
                        {
                            periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Lima_TipoMOD_Hogar.Clone();
                        }
                        else
                        {
                            periodo_Tipo_MOD_Temp = (double[,])periodo_Total_Ciudades_TipoMOD_Hogar.Clone();
                        }

                        for (int k = 0; k < periodo_Tipo_Temp_Unid.GetLength(0); k++) //FILAS
                        {
                            if (CodigoTipo == periodo_Tipo_Temp_Unid[k, 0].ToString())
                            {                                

                                t1 = double.Parse(reader_1[2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / periodo_Tipo_Temp_Unid[k, 1] * 100 : 0;
                                t2 = double.Parse(reader_1[3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / periodo_Tipo_Temp_Unid[k, 2] * 100 : 0;
                                t3 = double.Parse(reader_1[4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / periodo_Tipo_Temp_Unid[k, 3] * 100 : 0;
                                t4 = double.Parse(reader_1[5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / periodo_Tipo_Temp_Unid[k, 4] * 100 : 0;
                                t5 = double.Parse(reader_1[6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / periodo_Tipo_Temp_Unid[k, 5] * 100 : 0;
                                t6 = double.Parse(reader_1[7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / periodo_Tipo_Temp_Unid[k, 6] * 100 : 0;
                                t7 = double.Parse(reader_1[8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / periodo_Tipo_Temp_Unid[k, 7] * 100 : 0;
                                t8 = double.Parse(reader_1[9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / periodo_Tipo_Temp_Unid[k, 8] * 100 : 0;
                                t9 = double.Parse(reader_1[10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / periodo_Tipo_Temp_Unid[k, 9] * 100 : 0;
                                t10 = double.Parse(reader_1[11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / periodo_Tipo_Temp_Unid[k, 10] * 100 : 0;
                                t11 = double.Parse(reader_1[12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / periodo_Tipo_Temp_Unid[k, 11] * 100 : 0;
                                t12 = double.Parse(reader_1[13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / periodo_Tipo_Temp_Unid[k, 12] * 100 : 0;
                                t13 = double.Parse(reader_1[14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / periodo_Tipo_Temp_Unid[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }

                        // CALCULAR UNIDADES PROMEDIO (HOG.) MODALIDAD POR TIPOS"                        
                                               
                        for (int i = 0; i < periodo_Tipo_MOD_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoModalidad == periodo_Tipo_MOD_Temp[i, 0].ToString() && CodigoTipo == periodo_Tipo_MOD_Temp[i, 1].ToString())
                            {                                
                                t1 = double.Parse(periodo_Tipo_MOD_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Tipo_MOD_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Tipo_MOD_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Tipo_MOD_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Tipo_MOD_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Tipo_MOD_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Tipo_MOD_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Tipo_MOD_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Tipo_MOD_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Tipo_MOD_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Tipo_MOD_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Tipo_MOD_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Tipo_MOD_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Tipo_MOD_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Tipo_MOD_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Tipo_MOD_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Tipo_MOD_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Tipo_MOD_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Tipo_MOD_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Tipo_MOD_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Tipo_MOD_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Tipo_MOD_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Tipo_MOD_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Tipo_MOD_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Tipo_MOD_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Tipo_MOD_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1__, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11 , t12 , t13 );

                                Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado_, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                break;
                            }
                        }
                    }                   
                }
            }

            // CODIGO RECORRIDO POR TIPO Y NSE SHARE
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
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

                int columnas = 0;

                if (V1 == "Consolidado")
                {
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Mercado_Tipo_Unidad.Clone();
                    periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Mercado_TipoNSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Lima_Tipo_Unidad.Clone();
                    periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Lima_TipoNSE_Hogar.Clone();
                }
                else
                {
                    periodo_Tipo_Temp_Unid = (double[,])periodo_Total_Ciudades_Tipo_Unidad.Clone();
                    periodo_Tipo_NSE_Temp = (double[,])periodo_Total_Ciudades_TipoNSE_Hogar.Clone();
                }

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    while (reader_1.Read())
                    {
                        CodigoNSE = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_NSE.GetLength(0); i++) //  
                        {
                            if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                            {
                                V1_ = Codigo_Nombres_NSE[i, 1].Substring(3, Codigo_Nombres_NSE[i, 1].Length - 3); ;
                                Mercado = Codigo_Nombres_NSE[i, 1];
                                break;
                            }
                        }
                        CodigoTipo = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // 
                        {
                            if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                            {
                                V1__ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                                Mercado_ = Codigo_Nombres_Tipo[i, 1];
                                break;
                            }
                        }

                        for (int k = 0; k < periodo_Tipo_Temp_Unid.GetLength(0); k++) //FILAS
                        {
                            if (CodigoTipo == periodo_Tipo_Temp_Unid[k, 0].ToString())
                            {                        
                                t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Tipo_Temp_Unid[k, 1] * 100 : 0;
                                t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Tipo_Temp_Unid[k, 2] * 100 : 0;
                                t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Tipo_Temp_Unid[k, 3] * 100 : 0;
                                t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Tipo_Temp_Unid[k, 4] * 100 : 0;
                                t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Tipo_Temp_Unid[k, 5] * 100 : 0;
                                t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Tipo_Temp_Unid[k, 6] * 100 : 0;
                                t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Tipo_Temp_Unid[k, 7] * 100 : 0;
                                t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Tipo_Temp_Unid[k, 8] * 100 : 0;
                                t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Tipo_Temp_Unid[k, 9] * 100 : 0;
                                t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Tipo_Temp_Unid[k, 10] * 100 : 0;
                                t11 = reader_1.IsDBNull(12) ? 0: double.Parse(reader_1[12].ToString()) / periodo_Tipo_Temp_Unid[k, 11] * 100;
                                t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Tipo_Temp_Unid[k, 12] * 100 : 0;
                                t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Tipo_Temp_Unid[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(Mercado, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }

                        // CALCULAR UNIDADES PROMEDIO (HOG.) MODALIDAD POR TIPOS" 
                        for (int i = 0; i < periodo_Tipo_NSE_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoNSE == periodo_Tipo_NSE_Temp[i, 0].ToString() && CodigoTipo == periodo_Tipo_NSE_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Tipo_NSE_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Tipo_NSE_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Tipo_NSE_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Tipo_NSE_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Tipo_NSE_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Tipo_NSE_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Tipo_NSE_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Tipo_NSE_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Tipo_NSE_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Tipo_NSE_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Tipo_NSE_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Tipo_NSE_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Tipo_NSE_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Tipo_NSE_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Tipo_NSE_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Tipo_NSE_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Tipo_NSE_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Tipo_NSE_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Tipo_NSE_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Tipo_NSE_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Tipo_NSE_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Tipo_NSE_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Tipo_NSE_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Tipo_NSE_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Tipo_NSE_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Tipo_NSE_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(Mercado, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado_, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIA Y NSE SHARE
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_CATEGORIA_NSE"))
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

                int columnas = 0;

                if (V1 == "Consolidado")
                {
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Mercado_Categoria_Unidad.Clone();
                    periodo_Cate_NSE_Temp = (double[,])periodo_Total_Mercado_CateNSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Lima_Categoria_Unidad.Clone();
                    periodo_Cate_NSE_Temp = (double[,])periodo_Total_Lima_CateNSE_Hogar.Clone();
                }
                else
                {
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Ciudades_Categoria_Unidad.Clone();
                    periodo_Cate_NSE_Temp = (double[,])periodo_Total_Ciudades_CateNSE_Hogar.Clone();
                }

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    while (reader_1.Read())
                    {
                        CodigoNSE = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_NSE.GetLength(0); i++) //  
                        {
                            if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                            {
                                Mercado = Codigo_Nombres_NSE[i, 1].Substring(3, Codigo_Nombres_NSE[i, 1].Length - 3); ;
                                V1_ = Codigo_Nombres_NSE[i, 1];
                                break;
                            }
                        }
                        CodigoCategoria = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Categoria.GetLength(0); i++) // 
                        {
                            if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                            {
                                V1__ = Codigo_Nombres_Categoria[i, 1].Substring(4, Codigo_Nombres_Categoria[i, 1].Length - 4);
                                Mercado_ = Codigo_Nombres_Categoria[i, 1];
                                break;
                            }
                        }

                        for (int k = 0; k < periodo_Cat_Temp_Unid.GetLength(0); k++) //FILAS
                        {
                            if (CodigoCategoria == periodo_Cat_Temp_Unid[k, 0].ToString())
                            {
                                t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Cat_Temp_Unid[k, 1] * 100 : 0;
                                t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Cat_Temp_Unid[k, 2] * 100 : 0;
                                t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Cat_Temp_Unid[k, 3] * 100 : 0;
                                t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Cat_Temp_Unid[k, 4] * 100 : 0;
                                t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Cat_Temp_Unid[k, 5] * 100 : 0;
                                t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Cat_Temp_Unid[k, 6] * 100 : 0;
                                t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Cat_Temp_Unid[k, 7] * 100 : 0;
                                t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Cat_Temp_Unid[k, 8] * 100 : 0;
                                t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Cat_Temp_Unid[k, 9] * 100 : 0;
                                t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Cat_Temp_Unid[k, 10] * 100 : 0;
                                t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Cat_Temp_Unid[k, 11] * 100;
                                t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Cat_Temp_Unid[k, 12] * 100 : 0;
                                t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Cat_Temp_Unid[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }

                        // CALCULAR UNIDADES PROMEDIO (HOG.) MODALIDAD POR TIPOS" 
                        for (int i = 0; i < periodo_Cate_NSE_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoNSE == periodo_Cate_NSE_Temp[i, 0].ToString() && CodigoCategoria == periodo_Cate_NSE_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Cate_NSE_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Cate_NSE_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Cate_NSE_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Cate_NSE_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Cate_NSE_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Cate_NSE_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Cate_NSE_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Cate_NSE_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Cate_NSE_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Cate_NSE_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Cate_NSE_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Cate_NSE_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Cate_NSE_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Cate_NSE_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Cate_NSE_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Cate_NSE_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Cate_NSE_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Cate_NSE_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Cate_NSE_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Cate_NSE_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Cate_NSE_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Cate_NSE_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Cate_NSE_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Cate_NSE_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Cate_NSE_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Cate_NSE_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado_, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIA Y MODALIDAD SHARE
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_CATEGORIA_MODALIDAD"))
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

                int columnas = 0;

                if (V1 == "Consolidado")
                {
                    periodo_Canal_Temp_Unid = (double[,])periodo_Total_Mercado_Modalidad_Unidad.Clone();
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Mercado_Categoria_Unidad.Clone();                    
                    periodo_Cate_MOD_Temp = (double[,])periodo_Total_Mercado_CateMOD_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Canal_Temp_Unid = (double[,])periodo_Total_Lima_Modalidad_Unidad.Clone();
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Lima_Categoria_Unidad.Clone();
                    periodo_Cate_MOD_Temp = (double[,])periodo_Total_Lima_CateMOD_Hogar.Clone();
                }
                else
                {
                    periodo_Canal_Temp_Unid = (double[,])periodo_Total_Ciudades_Modalidad_Unidad.Clone();
                    periodo_Cat_Temp_Unid = (double[,])periodo_Total_Ciudades_Categoria_Unidad.Clone();
                    periodo_Cate_MOD_Temp = (double[,])periodo_Total_Ciudades_CateMOD_Hogar.Clone();
                }

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    while (reader_1.Read())
                    {
                        CodigoModalidad = reader_1[0].ToString();
                        for (int i = 0; i < Codigo_Nombres_Modalidad.GetLength(0); i++) //  
                        {
                            if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                            {
                                V1__ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                                Mercado = Codigo_Nombres_Modalidad[i, 1];
                                break;
                            }
                        }
                        CodigoCategoria = reader_1[1].ToString();
                        for (int i = 0; i < Codigo_Nombres_Categoria.GetLength(0); i++) // 
                        {
                            if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                            {
                                V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3);
                                Mercado_ = Codigo_Nombres_Categoria[i, 1];
                                break;
                            }
                        }
                        // SHARE CATEGORIA POR MODALIDAD SOBRE MODALIDAD
                        for (int k = 0; k < periodo_Canal_Temp_Unid.GetLength(0); k++) //FILAS
                        {
                            if (CodigoModalidad == periodo_Canal_Temp_Unid[k, 0].ToString())
                            {
                                t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Canal_Temp_Unid[k, 1] * 100 : 0;
                                t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Canal_Temp_Unid[k, 2] * 100 : 0;
                                t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Canal_Temp_Unid[k, 3] * 100 : 0;
                                t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Canal_Temp_Unid[k, 4] * 100 : 0;
                                t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Canal_Temp_Unid[k, 5] * 100 : 0;
                                t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Canal_Temp_Unid[k, 6] * 100 : 0;
                                t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Canal_Temp_Unid[k, 7] * 100 : 0;
                                t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Canal_Temp_Unid[k, 8] * 100 : 0;
                                t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Canal_Temp_Unid[k, 9] * 100 : 0;
                                t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Canal_Temp_Unid[k, 10] * 100 : 0;
                                t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Canal_Temp_Unid[k, 11] * 100;
                                t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Canal_Temp_Unid[k, 12] * 100 : 0;
                                t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Canal_Temp_Unid[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(V1_, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }
                        // SHARE CATEGORIA POR MODALIDAD SOBRE CATEGORIA
                        for (int k = 0; k < periodo_Cat_Temp_Unid.GetLength(0); k++) //FILAS
                        {
                            if (CodigoCategoria == periodo_Cat_Temp_Unid[k, 0].ToString())
                            {
                                t1 = reader_1[2] != DBNull.Value ? double.Parse(reader_1[2].ToString()) / periodo_Cat_Temp_Unid[k, 1] * 100 : 0;
                                t2 = reader_1[3] != DBNull.Value ? double.Parse(reader_1[3].ToString()) / periodo_Cat_Temp_Unid[k, 2] * 100 : 0;
                                t3 = reader_1[4] != DBNull.Value ? double.Parse(reader_1[4].ToString()) / periodo_Cat_Temp_Unid[k, 3] * 100 : 0;
                                t4 = reader_1[5] != DBNull.Value ? double.Parse(reader_1[5].ToString()) / periodo_Cat_Temp_Unid[k, 4] * 100 : 0;
                                t5 = reader_1[6] != DBNull.Value ? double.Parse(reader_1[6].ToString()) / periodo_Cat_Temp_Unid[k, 5] * 100 : 0;
                                t6 = reader_1[7] != DBNull.Value ? double.Parse(reader_1[7].ToString()) / periodo_Cat_Temp_Unid[k, 6] * 100 : 0;
                                t7 = reader_1[8] != DBNull.Value ? double.Parse(reader_1[8].ToString()) / periodo_Cat_Temp_Unid[k, 7] * 100 : 0;
                                t8 = reader_1[9] != DBNull.Value ? double.Parse(reader_1[9].ToString()) / periodo_Cat_Temp_Unid[k, 8] * 100 : 0;
                                t9 = reader_1[10] != DBNull.Value ? double.Parse(reader_1[10].ToString()) / periodo_Cat_Temp_Unid[k, 9] * 100 : 0;
                                t10 = reader_1[11] != DBNull.Value ? double.Parse(reader_1[11].ToString()) / periodo_Cat_Temp_Unid[k, 10] * 100 : 0;
                                t11 = reader_1.IsDBNull(12) ? 0 : double.Parse(reader_1[12].ToString()) / periodo_Cat_Temp_Unid[k, 11] * 100;
                                t12 = reader_1[13] != DBNull.Value ? double.Parse(reader_1[13].ToString()) / periodo_Cat_Temp_Unid[k, 12] * 100 : 0;
                                t13 = reader_1[14] != DBNull.Value ? double.Parse(reader_1[14].ToString()) / periodo_Cat_Temp_Unid[k, 13] * 100 : 0;

                                Actualizar_BD_Penetracion(V1__, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                                break;
                            }
                        }

                        // CALCULAR UNIDADES PROMEDIO (HOG.) MODALIDAD POR CATEGORIA" 
                        for (int i = 0; i < periodo_Cate_MOD_Temp.GetLength(0); i++) // GetLength(0) lee numero de filas array
                        {
                            if (CodigoModalidad == periodo_Cate_MOD_Temp[i, 0].ToString() && CodigoCategoria == periodo_Cate_MOD_Temp[i, 1].ToString())
                            {
                                t1 = double.Parse(periodo_Cate_MOD_Temp[i, 2].ToString()) != 0 ? double.Parse(reader_1[2].ToString()) / (periodo_Cate_MOD_Temp[i, 2] / 1000) : 0;
                                t2 = double.Parse(periodo_Cate_MOD_Temp[i, 3].ToString()) != 0 ? double.Parse(reader_1[3].ToString()) / (periodo_Cate_MOD_Temp[i, 3] / 1000) : 0;
                                t3 = double.Parse(periodo_Cate_MOD_Temp[i, 4].ToString()) != 0 ? double.Parse(reader_1[4].ToString()) / (periodo_Cate_MOD_Temp[i, 4] / 1000) : 0;
                                t4 = double.Parse(periodo_Cate_MOD_Temp[i, 5].ToString()) != 0 ? double.Parse(reader_1[5].ToString()) / (periodo_Cate_MOD_Temp[i, 5] / 1000) : 0;
                                t5 = double.Parse(periodo_Cate_MOD_Temp[i, 6].ToString()) != 0 ? double.Parse(reader_1[6].ToString()) / (periodo_Cate_MOD_Temp[i, 6] / 1000) : 0;
                                t6 = double.Parse(periodo_Cate_MOD_Temp[i, 7].ToString()) != 0 ? double.Parse(reader_1[7].ToString()) / (periodo_Cate_MOD_Temp[i, 7] / 1000) : 0;
                                t7 = double.Parse(periodo_Cate_MOD_Temp[i, 8].ToString()) != 0 ? double.Parse(reader_1[8].ToString()) / (periodo_Cate_MOD_Temp[i, 8] / 1000) : 0;
                                t8 = double.Parse(periodo_Cate_MOD_Temp[i, 9].ToString()) != 0 ? double.Parse(reader_1[9].ToString()) / (periodo_Cate_MOD_Temp[i, 9] / 1000) : 0;
                                t9 = double.Parse(periodo_Cate_MOD_Temp[i, 10].ToString()) != 0 ? double.Parse(reader_1[10].ToString()) / (periodo_Cate_MOD_Temp[i, 10] / 1000) : 0;
                                t10 = double.Parse(periodo_Cate_MOD_Temp[i, 11].ToString()) != 0 ? double.Parse(reader_1[11].ToString()) / (periodo_Cate_MOD_Temp[i, 11] / 1000) : 0;
                                t11 = double.Parse(periodo_Cate_MOD_Temp[i, 12].ToString()) != 0 ? double.Parse(reader_1[12].ToString()) / (periodo_Cate_MOD_Temp[i, 12] / 1000) : 0;
                                t12 = double.Parse(periodo_Cate_MOD_Temp[i, 13].ToString()) != 0 ? double.Parse(reader_1[13].ToString()) / (periodo_Cate_MOD_Temp[i, 13] / 1000) : 0;
                                t13 = double.Parse(periodo_Cate_MOD_Temp[i, 14].ToString()) != 0 ? double.Parse(reader_1[14].ToString()) / (periodo_Cate_MOD_Temp[i, 14] / 1000) : 0;

                                Actualizar_BD(V1_, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                Actualizar_BD(V1__, "Suma", "UNIDADES PROMEDIO (HOG.)", Ciudad_, Mercado_, "UNIDADES PROMEDIO (HOG.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                                break;
                            }
                        }
                    }
                }
            }

        }

        // HOGARES
        public void Universo_Periodos()
        {
            //RECUPERANDO Y ALMACENANDO UNIVERSO DE HOGARES DE LOS DIFERENTES PERIODOS
            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIVERSOS_PERIODOS_SELECT_NSE"))
            {
                int rows = 0;
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_Universos_NSE[rows, i] = valor_1;
                        }
                        rows++;
                    }
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIVERSOS_PERIODOS_SELECT"))
            {
                int rows = 0;
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            periodo_Universos[rows, i] = valor_1;
                        }
                        rows++;
                    }
                }
            }
        }
        public void Periodos_Cosmeticos_Total_Hogares(string xCiudad)
        {
            int Universo_Ciudad;
            Variable = "HOGARES";            
            V2 = "";

            if (xCiudad == "1")
            {
                V1 = "Lima";
                Ciudad_ = "1. Capital";
                V1_ = "Cosmeticos";
                Universo_Ciudad = 1;
            }
            else if (xCiudad == "1,2,5")
            {
                V1 = "Consolidado";
                Ciudad_ = "0. Consolidado";
                Universo_Ciudad = 3;
            }
            else
            {
                V1 = "Ciudades";
                Ciudad_ = "2. Ciudades";
                V1_ = "Cosmeticos";
                Universo_Ciudad = 2;
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                int rows = 0;
                int rowUniverso = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_Hogar[rows, i] = valor_1;
                                periodo_Temp[rows, i] = valor_1;
                                rowUniverso = 3;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Hogar[rows, i] = valor_1;
                                periodo_Temp[rows, i] = valor_1;
                                rowUniverso = 1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Hogar[rows, i] = valor_1;
                                periodo_Temp[rows, i] = valor_1;
                                rowUniverso = 2;
                            }
                        }
                        rows++;
                    }

                    Actualizar_BD(V1, "Suma", Variable, "0. Consolidado", "0. Cosmeticos", "HOGARES", "MENSUAL", periodo_Temp[0, 0], periodo_Temp[0, 1], periodo_Temp[0, 2], periodo_Temp[0, 3], periodo_Temp[0, 4], periodo_Temp[0, 5], periodo_Temp[0, 6], periodo_Temp[0, 7], periodo_Temp[0, 8], periodo_Temp[0, 9], periodo_Temp[0, 10], periodo_Temp[0, 11], periodo_Temp[0, 12]);

                    if (xCiudad != "1,2,5") //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS
                    {
                        Actualizar_BD("Cosmeticos", "Suma", Variable, Ciudad_, "0. Cosmeticos", "HOGARES", "MENSUAL", periodo_Temp[0, 0], periodo_Temp[0, 1], periodo_Temp[0, 2], periodo_Temp[0, 3], periodo_Temp[0, 4], periodo_Temp[0, 5], periodo_Temp[0, 6], periodo_Temp[0, 7], periodo_Temp[0, 8], periodo_Temp[0, 9], periodo_Temp[0, 10], periodo_Temp[0, 11], periodo_Temp[0, 12]);
                    }

                    // CALCULAR PENETRACION POR HOGAR                                    
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    for (int i = 0; i < periodo_Universos.GetLength(0); i++) // GetLength(0) lee numero de filas array
                    {
                        if (rowUniverso == periodo_Universos[i,0])
                        {                            
                            t1 = double.Parse(periodo_Universos[i, 1].ToString()) != 0 ? periodo_Temp[0, 0] / periodo_Universos[i, 1] * 100 : 0;
                            t2 = double.Parse(periodo_Universos[i, 2].ToString()) != 0 ? periodo_Temp[0, 1] / periodo_Universos[i, 2] * 100 : 0;
                            t3 = double.Parse(periodo_Universos[i, 3].ToString()) != 0 ? periodo_Temp[0, 2] / periodo_Universos[i, 3] * 100 : 0;
                            t4 = double.Parse(periodo_Universos[i, 4].ToString()) != 0 ? periodo_Temp[0, 3] / periodo_Universos[i, 4] * 100 : 0;
                            t5 = double.Parse(periodo_Universos[i, 5].ToString()) != 0 ? periodo_Temp[0, 4] / periodo_Universos[i, 5] * 100 : 0;
                            t6 = double.Parse(periodo_Universos[i, 6].ToString()) != 0 ? periodo_Temp[0, 5] / periodo_Universos[i, 6] * 100 : 0;
                            t7 = double.Parse(periodo_Universos[i, 7].ToString()) != 0 ? periodo_Temp[0, 6] / periodo_Universos[i, 7] * 100 : 0;
                            t8 = double.Parse(periodo_Universos[i, 8].ToString()) != 0 ? periodo_Temp[0, 7] / periodo_Universos[i, 8] * 100 : 0;
                            t9 = double.Parse(periodo_Universos[i, 9].ToString()) != 0 ? periodo_Temp[0, 8] / periodo_Universos[i, 9] * 100 : 0;
                            t10 = double.Parse(periodo_Universos[i, 10].ToString()) != 0 ? periodo_Temp[0, 9] / periodo_Universos[i, 10] * 100 : 0;
                            t11 = double.Parse(periodo_Universos[i, 11].ToString()) != 0 ? periodo_Temp[0, 10] / periodo_Universos[i, 11] * 100 : 0;
                            t12 = double.Parse(periodo_Universos[i, 12].ToString()) != 0 ? periodo_Temp[0, 11] / periodo_Universos[i, 12] * 100 : 0;
                            t13 = double.Parse(periodo_Universos[i, 13].ToString()) != 0 ? periodo_Temp[0, 12] / periodo_Universos[i, 13] * 100 : 0;

                            Actualizar_BD_Penetracion(V1, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                            if (rowUniverso != 3) //SOLO PARA CREAR LOS REGISTROS V1 = COSMETICOS
                            {
                                Actualizar_BD("Cosmeticos", "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                            }
                            break;
                        }
                    }
                }
            }

            // CODIGO RECORRIDO POR NSE              
            double[,] periodo_NSE_Temp_Unidad = new double[5, 14]; //CREAR LA MATRIZ PARA BORRAR VALORES PASADOS
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                int rows = 0;
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_NSE_Hogar[rows, i ] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i ] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_NSE_Hogar[rows, i ] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i ] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_NSE_Hogar[rows, i] = valor_1;
                                periodo_NSE_Temp_Unidad[rows, i ] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
           
            string CodigoNSE;
            for (int k = 0; k < periodo_NSE_Temp_Unidad.GetLength(0); k++) //FILAS
            {
                CodigoNSE = periodo_NSE_Temp_Unidad[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        break;
                        //Mercado = Codigo_Nombres_NSE[i, 1];
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos", "HOGARES", "MENSUAL", periodo_NSE_Temp_Unidad[k, 1], periodo_NSE_Temp_Unidad[k, 2], periodo_NSE_Temp_Unidad[k, 3], periodo_NSE_Temp_Unidad[k, 4], periodo_NSE_Temp_Unidad[k, 5], periodo_NSE_Temp_Unidad[k, 6], periodo_NSE_Temp_Unidad[k, 7], periodo_NSE_Temp_Unidad[k, 8], periodo_NSE_Temp_Unidad[k, 9], periodo_NSE_Temp_Unidad[k, 10], periodo_NSE_Temp_Unidad[k, 11], periodo_NSE_Temp_Unidad[k, 12], periodo_NSE_Temp_Unidad[k, 13]);

                // CALCULAR PENETRACION POR HOGAR                
                for (int i = 0; i < periodo_Universos_NSE.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (Universo_Ciudad == periodo_Universos_NSE[i, 0] && CodigoNSE == periodo_Universos_NSE[i, 1].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = double.Parse(periodo_Universos_NSE[i, 2].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 1] / periodo_Universos_NSE[i, 2] * 100 : 0;
                        t2 = double.Parse(periodo_Universos_NSE[i, 3].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 2] / periodo_Universos_NSE[i, 3] * 100 : 0;
                        t3 = double.Parse(periodo_Universos_NSE[i, 4].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 3] / periodo_Universos_NSE[i, 4] * 100 : 0;
                        t4 = double.Parse(periodo_Universos_NSE[i, 5].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 4] / periodo_Universos_NSE[i, 5] * 100 : 0;
                        t5 = double.Parse(periodo_Universos_NSE[i, 6].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 5] / periodo_Universos_NSE[i, 6] * 100 : 0;
                        t6 = double.Parse(periodo_Universos_NSE[i, 7].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 6] / periodo_Universos_NSE[i, 7] * 100 : 0;
                        t7 = double.Parse(periodo_Universos_NSE[i, 8].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 7] / periodo_Universos_NSE[i, 8] * 100 : 0;
                        t8 = double.Parse(periodo_Universos_NSE[i, 9].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 8] / periodo_Universos_NSE[i, 9] * 100 : 0;
                        t9 = double.Parse(periodo_Universos_NSE[i, 10].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 9] / periodo_Universos_NSE[i, 10] * 100 : 0;
                        t10 = double.Parse(periodo_Universos_NSE[i, 11].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 10] / periodo_Universos_NSE[i, 11] * 100 : 0;
                        t11 = double.Parse(periodo_Universos_NSE[i, 12].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 11] / periodo_Universos_NSE[i, 12] * 100 : 0;
                        t12 = double.Parse(periodo_Universos_NSE[i, 13].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 12] / periodo_Universos_NSE[i, 13] * 100 : 0;
                        t13 = double.Parse(periodo_Universos_NSE[i, 14].ToString()) != 0 ? periodo_NSE_Temp_Unidad[k, 13] / periodo_Universos_NSE[i, 14] * 100 : 0;

                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        break;
                    }
                }

            }

            // CODIGO RECORRIDO POR CATEGORIAS           
            string CodigoCategoria;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_CATEGORIA"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                int rows = 0;
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_Categoria_Hogar[rows, i ] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i ] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_Hogar[rows, i ] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i ] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_Hogar[rows, i ] = valor_1;
                                periodo_Cat_Temp_Unid[rows, i ] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int k = 0; k < periodo_Cat_Temp_Unid.GetLength(0); k++) //FILAS
            {
                CodigoCategoria = periodo_Cat_Temp_Unid[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Categoria[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Cat_Temp_Unid[k, 1], periodo_Cat_Temp_Unid[k, 2], periodo_Cat_Temp_Unid[k, 3], periodo_Cat_Temp_Unid[k, 4], periodo_Cat_Temp_Unid[k, 5], periodo_Cat_Temp_Unid[k, 6], periodo_Cat_Temp_Unid[k, 7], periodo_Cat_Temp_Unid[k, 8], periodo_Cat_Temp_Unid[k, 9], periodo_Cat_Temp_Unid[k, 10], periodo_Cat_Temp_Unid[k, 11], periodo_Cat_Temp_Unid[k, 12], periodo_Cat_Temp_Unid[k, 13]);

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos2", "HOGARES", "MENSUAL", periodo_Cat_Temp_Unid[k, 1], periodo_Cat_Temp_Unid[k, 2], periodo_Cat_Temp_Unid[k, 3], periodo_Cat_Temp_Unid[k, 4], periodo_Cat_Temp_Unid[k, 5], periodo_Cat_Temp_Unid[k, 6], periodo_Cat_Temp_Unid[k, 7], periodo_Cat_Temp_Unid[k, 8], periodo_Cat_Temp_Unid[k, 9], periodo_Cat_Temp_Unid[k, 10], periodo_Cat_Temp_Unid[k, 11], periodo_Cat_Temp_Unid[k, 12], periodo_Cat_Temp_Unid[k, 13]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Hogar.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Total_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                    t1 = double.Parse(periodo_Total_Temporal[i, 0].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 1] / periodo_Total_Temporal[i, 0] * 100 : 0;
                    t2 = double.Parse(periodo_Total_Temporal[i, 1].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 2] / periodo_Total_Temporal[i, 1] * 100 : 0;
                    t3 = double.Parse(periodo_Total_Temporal[i, 2].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 3] / periodo_Total_Temporal[i, 2] * 100 : 0;
                    t4 = double.Parse(periodo_Total_Temporal[i, 3].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 4] / periodo_Total_Temporal[i, 3] * 100 : 0;
                    t5 = double.Parse(periodo_Total_Temporal[i, 4].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 5] / periodo_Total_Temporal[i, 4] * 100 : 0;
                    t6 = double.Parse(periodo_Total_Temporal[i, 5].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 6] / periodo_Total_Temporal[i, 5] * 100 : 0;
                    t7 = double.Parse(periodo_Total_Temporal[i, 6].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 7] / periodo_Total_Temporal[i, 6] * 100 : 0;
                    t8 = double.Parse(periodo_Total_Temporal[i, 7].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 8] / periodo_Total_Temporal[i, 7] * 100 : 0;
                    t9 = double.Parse(periodo_Total_Temporal[i, 8].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 9] / periodo_Total_Temporal[i, 8] * 100 : 0;
                    t10 = double.Parse(periodo_Total_Temporal[i, 9].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 10] / periodo_Total_Temporal[i, 9] * 100 : 0;
                    t11 = double.Parse(periodo_Total_Temporal[i, 10].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 11] / periodo_Total_Temporal[i, 10] * 100 : 0;
                    t12 = double.Parse(periodo_Total_Temporal[i, 11].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 12] / periodo_Total_Temporal[i, 11] * 100 : 0;
                    t13 = double.Parse(periodo_Total_Temporal[i, 12].ToString()) != 0 ? periodo_Cat_Temp_Unid[k, 13] / periodo_Total_Temporal[i, 12] * 100 : 0;

                    Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos2", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    
                }
            }

            // CODIGO RECORRIDO POR TIPOS      
            double[,] periodo_Tipo_Temp_Unid = new double[15, 14];  // limpiando la matriz de datos historicos
            double[,] periodo_Total_Mercado_Tipo_Hogar = new double[15, 14];  // LIMA POR  UNIDAD
            double[,] periodo_Total_Lima_Tipo_Hogar = new double[15, 14];  // 
            double[,] periodo_Total_Ciudades_Tipo_Hogar = new double[15, 14];  // 
            string CodigoTipo;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_TIPO"))
            {                
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                int rows = 0;
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_Tipo_Hogar[rows, i ] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i ] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Tipo_Hogar[rows, i ] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i ] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Tipo_Hogar[rows, i ] = valor_1;
                                periodo_Tipo_Temp_Unid[rows, i ] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Tipo_Temp_Unid.GetLength(0); ii++) // GetLength(0) lee numero de filas array
            {
                CodigoTipo = periodo_Tipo_Temp_Unid[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.GetLength(0); i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                    {
                        V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado = Codigo_Nombres_Tipo[i, 1];
                        break;
                    }
                }              

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Tipo_Temp_Unid[ii, 1], periodo_Tipo_Temp_Unid[ii, 2], periodo_Tipo_Temp_Unid[ii, 3], periodo_Tipo_Temp_Unid[ii, 4], periodo_Tipo_Temp_Unid[ii, 5], periodo_Tipo_Temp_Unid[ii, 6], periodo_Tipo_Temp_Unid[ii, 7], periodo_Tipo_Temp_Unid[ii, 8], periodo_Tipo_Temp_Unid[ii, 9], periodo_Tipo_Temp_Unid[ii, 10], periodo_Tipo_Temp_Unid[ii, 11], periodo_Tipo_Temp_Unid[ii, 12], periodo_Tipo_Temp_Unid[ii, 13]);
                
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos2", "HOGARES", "MENSUAL", periodo_Tipo_Temp_Unid[ii, 1], periodo_Tipo_Temp_Unid[ii, 2], periodo_Tipo_Temp_Unid[ii, 3], periodo_Tipo_Temp_Unid[ii, 4], periodo_Tipo_Temp_Unid[ii, 5], periodo_Tipo_Temp_Unid[ii, 6], periodo_Tipo_Temp_Unid[ii, 7], periodo_Tipo_Temp_Unid[ii, 8], periodo_Tipo_Temp_Unid[ii, 9], periodo_Tipo_Temp_Unid[ii, 10], periodo_Tipo_Temp_Unid[ii, 11], periodo_Tipo_Temp_Unid[ii, 12], periodo_Tipo_Temp_Unid[ii, 13]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_Cat_Temporal = (double[,]) periodo_Total_Mercado_Categoria_Hogar.Clone();
                    periodo_Total_Mercado_Tipo_Hogar_ = (double[,])periodo_Total_Mercado_Tipo_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Cat_Temporal = (double[,]) periodo_Total_Lima_Categoria_Hogar.Clone();
                    periodo_Total_Lima_Tipo_Hogar_ = (double[,])periodo_Total_Lima_Tipo_Hogar.Clone();
                }
                else
                {
                    periodo_Cat_Temporal = (double[,]) periodo_Total_Ciudades_Categoria_Hogar.Clone();
                    periodo_Total_Ciudades_Tipo_Hogar_ = (double[,])periodo_Total_Ciudades_Tipo_Hogar.Clone();
                }

                int xCateg = 0;
                for (int i = 0; i < Codigo_Tipos_Categorias.GetLength(0); i++) 
                {
                    if (CodigoTipo == Codigo_Tipos_Categorias[i,0].ToString())
                    {
                        xCateg = Codigo_Tipos_Categorias[i,1];
                        break;
                    }
                }
                    
                for (int i = 0; i < periodo_Cat_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (xCateg == int.Parse(periodo_Cat_Temporal[i, 0].ToString()))
                    {
                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", 
                            periodo_Tipo_Temp_Unid[ii, 1] / periodo_Cat_Temporal[i, 1] * 100, 
                            periodo_Tipo_Temp_Unid[ii, 2] / periodo_Cat_Temporal[i, 2] * 100,
                            periodo_Tipo_Temp_Unid[ii, 3] / periodo_Cat_Temporal[i, 3] * 100,  
                            periodo_Tipo_Temp_Unid[ii, 4] / periodo_Cat_Temporal[i, 4] * 100,  
                            periodo_Tipo_Temp_Unid[ii, 5] / periodo_Cat_Temporal[i, 5] * 100,  
                            periodo_Tipo_Temp_Unid[ii, 6] / periodo_Cat_Temporal[i, 6] * 100,  
                            periodo_Tipo_Temp_Unid[ii, 7] / periodo_Cat_Temporal[i, 7] * 100,
                            periodo_Tipo_Temp_Unid[ii, 8] / periodo_Cat_Temporal[i, 8] * 100,
                            periodo_Tipo_Temp_Unid[ii, 9] / periodo_Cat_Temporal[i, 9] * 100,
                            periodo_Tipo_Temp_Unid[ii, 10] / periodo_Cat_Temporal[i, 10] * 100,
                            periodo_Tipo_Temp_Unid[ii, 11] / periodo_Cat_Temporal[i, 11] * 100,
                            periodo_Tipo_Temp_Unid[ii, 12] / periodo_Cat_Temporal[i, 12] * 100,
                            periodo_Tipo_Temp_Unid[ii, 13] / periodo_Cat_Temporal[i, 13] * 100);
                        //0. Cosmeticos2
                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos2", "PENETRACIONES (%)", "MENSUAL",
                           periodo_Tipo_Temp_Unid[ii, 1] / periodo_Cat_Temporal[i, 1] * 100,
                           periodo_Tipo_Temp_Unid[ii, 2] / periodo_Cat_Temporal[i, 2] * 100,
                           periodo_Tipo_Temp_Unid[ii, 3] / periodo_Cat_Temporal[i, 3] * 100,
                           periodo_Tipo_Temp_Unid[ii, 4] / periodo_Cat_Temporal[i, 4] * 100,
                           periodo_Tipo_Temp_Unid[ii, 5] / periodo_Cat_Temporal[i, 5] * 100,
                           periodo_Tipo_Temp_Unid[ii, 6] / periodo_Cat_Temporal[i, 6] * 100,
                           periodo_Tipo_Temp_Unid[ii, 7] / periodo_Cat_Temporal[i, 7] * 100,
                           periodo_Tipo_Temp_Unid[ii, 8] / periodo_Cat_Temporal[i, 8] * 100,
                           periodo_Tipo_Temp_Unid[ii, 9] / periodo_Cat_Temporal[i, 9] * 100,
                           periodo_Tipo_Temp_Unid[ii, 10] / periodo_Cat_Temporal[i, 10] * 100,
                           periodo_Tipo_Temp_Unid[ii, 11] / periodo_Cat_Temporal[i, 11] * 100,
                           periodo_Tipo_Temp_Unid[ii, 12] / periodo_Cat_Temporal[i, 12] * 100,
                           periodo_Tipo_Temp_Unid[ii, 13] / periodo_Cat_Temporal[i, 13] * 100);

                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR MODALIDAD      
            string CodigoModalidad;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MODALIDAD", DbType.String, Codigo_cadena_Modalidad);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                int rows = 0;
                int columnas = 0;
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_Modalidad_Hogar[rows, i] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Modalidad_Hogar[rows, i] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Modalidad_Hogar[rows, i] = valor_1;
                                periodo_Canal_Temp_Unid[rows, i] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int k = 0; k < periodo_Canal_Temp_Unid.GetLength(0); k++) //FILAS
            {
                CodigoModalidad = periodo_Canal_Temp_Unid[k, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                    {
                        V1_ = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Modalidad[i, 1];
                        break;
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Canal_Temp_Unid[k, 1], periodo_Canal_Temp_Unid[k, 2], periodo_Canal_Temp_Unid[k, 3], periodo_Canal_Temp_Unid[k, 4], periodo_Canal_Temp_Unid[k, 5], periodo_Canal_Temp_Unid[k, 6], periodo_Canal_Temp_Unid[k, 7], periodo_Canal_Temp_Unid[k, 8], periodo_Canal_Temp_Unid[k, 9], periodo_Canal_Temp_Unid[k, 10], periodo_Canal_Temp_Unid[k, 11], periodo_Canal_Temp_Unid[k, 12], periodo_Canal_Temp_Unid[k, 13]);

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, "0. Cosmeticos", "HOGARES", "MENSUAL", periodo_Canal_Temp_Unid[k, 1], periodo_Canal_Temp_Unid[k, 2], periodo_Canal_Temp_Unid[k, 3], periodo_Canal_Temp_Unid[k, 4], periodo_Canal_Temp_Unid[k, 5], periodo_Canal_Temp_Unid[k, 6], periodo_Canal_Temp_Unid[k, 7], periodo_Canal_Temp_Unid[k, 8], periodo_Canal_Temp_Unid[k, 9], periodo_Canal_Temp_Unid[k, 10], periodo_Canal_Temp_Unid[k, 11], periodo_Canal_Temp_Unid[k, 12], periodo_Canal_Temp_Unid[k, 13]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Mercado_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Lima_Hogar.Clone();
                }
                else
                {
                    periodo_Total_Temporal = (double[,])periodo_Total_Ciudades_Hogar.Clone();
                }

                //for (int i = 0; i < periodo_Total_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                //{
                double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                t1 = periodo_Total_Temporal[0, 0] != 0 ? periodo_Canal_Temp_Unid[k, 1] / periodo_Total_Temporal[0, 0] * 100 : 0;
                t2 = periodo_Total_Temporal[0, 1] != 0 ? periodo_Canal_Temp_Unid[k, 2] / periodo_Total_Temporal[0, 1] * 100 : 0;
                t3 = periodo_Total_Temporal[0, 2] != 0 ? periodo_Canal_Temp_Unid[k, 3] / periodo_Total_Temporal[0, 2] * 100 : 0;
                t4 = periodo_Total_Temporal[0, 3] != 0 ? periodo_Canal_Temp_Unid[k, 4] / periodo_Total_Temporal[0, 3] * 100 : 0;
                t5 = periodo_Total_Temporal[0, 4] != 0 ? periodo_Canal_Temp_Unid[k, 5] / periodo_Total_Temporal[0, 4] * 100 : 0;
                t6 = periodo_Total_Temporal[0, 5] != 0 ? periodo_Canal_Temp_Unid[k, 6] / periodo_Total_Temporal[0, 5] * 100 : 0;
                t7 = periodo_Total_Temporal[0, 6] != 0 ? periodo_Canal_Temp_Unid[k, 7] / periodo_Total_Temporal[0, 6] * 100 : 0;
                t8 = periodo_Total_Temporal[0, 7] != 0 ? periodo_Canal_Temp_Unid[k, 8] / periodo_Total_Temporal[0, 7] * 100 : 0;
                t9 = periodo_Total_Temporal[0, 8] != 0 ? periodo_Canal_Temp_Unid[k, 9] / periodo_Total_Temporal[0, 8] * 100 : 0;
                t10 = periodo_Total_Temporal[0, 9] != 0 ? periodo_Canal_Temp_Unid[k, 10] / periodo_Total_Temporal[0, 9] * 100 : 0;
                t11 = periodo_Total_Temporal[0, 10] != 0 ? periodo_Canal_Temp_Unid[k, 11] / periodo_Total_Temporal[0, 10] * 100 : 0;
                t12 = periodo_Total_Temporal[0, 11] != 0 ? periodo_Canal_Temp_Unid[k, 12] / periodo_Total_Temporal[0, 11] * 100 : 0;
                t13 = periodo_Total_Temporal[0, 12] != 0 ? periodo_Canal_Temp_Unid[k, 13] / periodo_Total_Temporal[0, 12] * 100 : 0;

                Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                //}
            }

            // CODIGO RECORRIDO POR TIPOS Y NSE   
            double[,] periodo_Tipo_NSE_Temp = new double[75, 15];  // 
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_TIPO_NSE"))
            {                
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                int rows = 0;
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_TipoNSE_Hogar[rows, i] = valor_1;
                                periodo_Tipo_NSE_Temp[rows, i] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_TipoNSE_Hogar[rows, i] = valor_1;
                                periodo_Tipo_NSE_Temp[rows, i] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_TipoNSE_Hogar[rows, i] = valor_1;
                                periodo_Tipo_NSE_Temp[rows, i] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Tipo_NSE_Temp.GetLength(0); ii++) // FILAS
            {
                CodigoTipo = periodo_Tipo_NSE_Temp[ii, 1].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE/2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                    {
                        //V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado = Codigo_Nombres_Tipo[i, 1];
                        break;
                    }
                }

                CodigoNSE = periodo_Tipo_NSE_Temp[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE/2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        break;
                        //Mercado = Codigo_Nombres_NSE[i, 1];
                    }
                }

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Tipo_NSE_Temp[ii, 2], periodo_Tipo_NSE_Temp[ii, 3], periodo_Tipo_NSE_Temp[ii, 4], periodo_Tipo_NSE_Temp[ii, 5], periodo_Tipo_NSE_Temp[ii, 6], periodo_Tipo_NSE_Temp[ii, 7], periodo_Tipo_NSE_Temp[ii, 8], periodo_Tipo_NSE_Temp[ii, 9], periodo_Tipo_NSE_Temp[ii, 10], periodo_Tipo_NSE_Temp[ii, 11], periodo_Tipo_NSE_Temp[ii, 12], periodo_Tipo_NSE_Temp[ii, 13], periodo_Tipo_NSE_Temp[ii, 14]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Mercado_NSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Lima_NSE_Hogar.Clone();
                }
                else
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Ciudades_NSE_Hogar.Clone();
                }

                for (int i = 0; i < periodo_NSE_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {                    
                    if (CodigoNSE == periodo_NSE_Temporal[i, 0].ToString())
                    {
                        double t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13;
                        t1 = periodo_NSE_Temporal[i, 1] != 0 ? periodo_Tipo_NSE_Temp[ii, 2] / periodo_NSE_Temporal[i, 1] * 100 : 0;
                        t2 = periodo_NSE_Temporal[i, 2] != 0 ? periodo_Tipo_NSE_Temp[ii, 3] / periodo_NSE_Temporal[i, 2] * 100 : 0;
                        t3 = periodo_NSE_Temporal[i, 3] != 0 ? periodo_Tipo_NSE_Temp[ii, 4] / periodo_NSE_Temporal[i, 3] * 100 : 0;
                        t4 = periodo_NSE_Temporal[i, 4] != 0 ? periodo_Tipo_NSE_Temp[ii, 5] / periodo_NSE_Temporal[i, 4] * 100 : 0;
                        t5 = periodo_NSE_Temporal[i, 5] != 0 ? periodo_Tipo_NSE_Temp[ii, 6] / periodo_NSE_Temporal[i, 5] * 100 : 0;
                        t6 = periodo_NSE_Temporal[i, 6] != 0 ? periodo_Tipo_NSE_Temp[ii, 7] / periodo_NSE_Temporal[i, 6] * 100 : 0;
                        t7 = periodo_NSE_Temporal[i, 7] != 0 ? periodo_Tipo_NSE_Temp[ii, 8] / periodo_NSE_Temporal[i, 7] * 100 : 0;
                        t8 = periodo_NSE_Temporal[i, 8] != 0 ? periodo_Tipo_NSE_Temp[ii, 9] / periodo_NSE_Temporal[i, 8] * 100 : 0;
                        t9 = periodo_NSE_Temporal[i, 9] != 0 ? periodo_Tipo_NSE_Temp[ii, 10] / periodo_NSE_Temporal[i, 9] * 100 : 0;
                        t10 = periodo_NSE_Temporal[i, 10] != 0 ? periodo_Tipo_NSE_Temp[ii, 11] / periodo_NSE_Temporal[i, 10] * 100 : 0;
                        t11 = periodo_NSE_Temporal[i, 11] != 0 ? periodo_Tipo_NSE_Temp[ii, 12] / periodo_NSE_Temporal[i, 11] * 100 : 0;
                        t12 = periodo_NSE_Temporal[i, 12] != 0 ? periodo_Tipo_NSE_Temp[ii, 13] / periodo_NSE_Temporal[i, 12] * 100 : 0;
                        t13 = periodo_NSE_Temporal[i, 13] != 0 ? periodo_Tipo_NSE_Temp[ii, 14] / periodo_NSE_Temporal[i, 13] * 100 : 0;

                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        break;
                    }
                }

            }

            // CODIGO RECORRIDO POR CATEGORIA Y NSE                              
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_CATEGORIA_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                int rows = 0;
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_CateNSE_Hogar[rows, i] = valor_1;
                                periodo_Cate_NSE_Temp[rows, i] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_CateNSE_Hogar[rows, i] = valor_1;
                                periodo_Cate_NSE_Temp[rows, i] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_CateNSE_Hogar[rows, i] = valor_1;
                                periodo_Cate_NSE_Temp[rows, i] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Cate_NSE_Temp.GetLength(0); ii++) // FILAS
            {
                CodigoCategoria = periodo_Cate_NSE_Temp[ii, 1].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        //V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Categoria[i, 1];
                        break;
                    }
                }

                CodigoNSE = periodo_Cate_NSE_Temp[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_NSE.Length / 2; i++) // SE DIVIDE/2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoNSE == Codigo_Nombres_NSE[i, 0])
                    {
                        V1_ = Codigo_Nombres_NSE[i, 1];
                        break;
                        //Mercado = Codigo_Nombres_NSE[i, 1];
                    }
                }
                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Cate_NSE_Temp[ii, 2], periodo_Cate_NSE_Temp[ii, 3], periodo_Cate_NSE_Temp[ii, 4], periodo_Cate_NSE_Temp[ii, 5], periodo_Cate_NSE_Temp[ii, 6], periodo_Cate_NSE_Temp[ii, 7], periodo_Cate_NSE_Temp[ii, 8], periodo_Cate_NSE_Temp[ii, 9], periodo_Cate_NSE_Temp[ii, 10], periodo_Cate_NSE_Temp[ii, 11], periodo_Cate_NSE_Temp[ii, 12], periodo_Cate_NSE_Temp[ii, 13], periodo_Cate_NSE_Temp[ii, 14]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Mercado_NSE_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Lima_NSE_Hogar.Clone();
                }
                else
                {
                    periodo_NSE_Temporal = (double[,])periodo_Total_Ciudades_NSE_Hogar.Clone();
                }

                for (int i = 0; i < periodo_NSE_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoNSE == periodo_NSE_Temporal[i, 0].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = periodo_NSE_Temporal[i, 1] != 0 ? periodo_Cate_NSE_Temp[ii, 2] / periodo_NSE_Temporal[i, 1] * 100 : 0;
                        t2 = periodo_NSE_Temporal[i, 2] != 0 ? periodo_Cate_NSE_Temp[ii, 3] / periodo_NSE_Temporal[i, 2] * 100 : 0;
                        t3 = periodo_NSE_Temporal[i, 3] != 0 ? periodo_Cate_NSE_Temp[ii, 4] / periodo_NSE_Temporal[i, 3] * 100 : 0;
                        t4 = periodo_NSE_Temporal[i, 4] != 0 ? periodo_Cate_NSE_Temp[ii, 5] / periodo_NSE_Temporal[i, 4] * 100 : 0;
                        t5 = periodo_NSE_Temporal[i, 5] != 0 ? periodo_Cate_NSE_Temp[ii, 6] / periodo_NSE_Temporal[i, 5] * 100 : 0;
                        t6 = periodo_NSE_Temporal[i, 6] != 0 ? periodo_Cate_NSE_Temp[ii, 7] / periodo_NSE_Temporal[i, 6] * 100 : 0;
                        t7 = periodo_NSE_Temporal[i, 7] != 0 ? periodo_Cate_NSE_Temp[ii, 8] / periodo_NSE_Temporal[i, 7] * 100 : 0;
                        t8 = periodo_NSE_Temporal[i, 8] != 0 ? periodo_Cate_NSE_Temp[ii, 9] / periodo_NSE_Temporal[i, 8] * 100 : 0;
                        t9 = periodo_NSE_Temporal[i, 9] != 0 ? periodo_Cate_NSE_Temp[ii, 10] / periodo_NSE_Temporal[i, 9] * 100 : 0;
                        t10 = periodo_NSE_Temporal[i, 10] != 0 ? periodo_Cate_NSE_Temp[ii, 11] / periodo_NSE_Temporal[i, 10] * 100 : 0;
                        t11 = periodo_NSE_Temporal[i, 11] != 0 ? periodo_Cate_NSE_Temp[ii, 12] / periodo_NSE_Temporal[i, 11] * 100 : 0;
                        t12 = periodo_NSE_Temporal[i, 12] != 0 ? periodo_Cate_NSE_Temp[ii, 13] / periodo_NSE_Temporal[i, 12] * 100 : 0;
                        t13 = periodo_NSE_Temporal[i, 13] != 0 ? periodo_Cate_NSE_Temp[ii, 14] / periodo_NSE_Temporal[i, 13] * 100 : 0;

                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR TIPOS Y MODALIDAD                              
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_TIPO_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                int rows = 0;
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_TipoMOD_Hogar[rows, i] = valor_1;
                                periodo_Tipo_MOD_Temp[rows, i] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_TipoMOD_Hogar[rows, i] = valor_1;
                                periodo_Tipo_MOD_Temp[rows, i] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_TipoMOD_Hogar[rows, i] = valor_1;
                                periodo_Tipo_MOD_Temp[rows, i] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Tipo_MOD_Temp.GetLength(0); ii++) // FILAS
            {
                CodigoTipo = periodo_Tipo_MOD_Temp[ii, 1].ToString();
                for (int i = 0; i < Codigo_Nombres_Tipo.Length / 2; i++) // SE DIVIDE/2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoTipo == Codigo_Nombres_Tipo[i, 0])
                    {
                        V1_ = Codigo_Nombres_Tipo[i, 1].Substring(4, Codigo_Nombres_Tipo[i, 1].Length - 4);
                        Mercado_ = Codigo_Nombres_Tipo[i, 1];
                        break;
                    }
                }

                CodigoModalidad = periodo_Tipo_MOD_Temp[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                    {
                        V2 = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); 
                        Mercado = Codigo_Nombres_Modalidad[i, 1];
                        break;
                    }
                }

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Tipo_MOD_Temp[ii, 2], periodo_Tipo_MOD_Temp[ii, 3], periodo_Tipo_MOD_Temp[ii, 4], periodo_Tipo_MOD_Temp[ii, 5], periodo_Tipo_MOD_Temp[ii, 6], periodo_Tipo_MOD_Temp[ii, 7], periodo_Tipo_MOD_Temp[ii, 8], periodo_Tipo_MOD_Temp[ii, 9], periodo_Tipo_MOD_Temp[ii, 10], periodo_Tipo_MOD_Temp[ii, 11], periodo_Tipo_MOD_Temp[ii, 12], periodo_Tipo_MOD_Temp[ii, 13], periodo_Tipo_MOD_Temp[ii, 14]);
                // MODALIDAD TIPO
                Actualizar_BD(V2, "Suma", Variable, Ciudad_, Mercado_, "HOGARES", "MENSUAL", periodo_Tipo_MOD_Temp[ii, 2], periodo_Tipo_MOD_Temp[ii, 3], periodo_Tipo_MOD_Temp[ii, 4], periodo_Tipo_MOD_Temp[ii, 5], periodo_Tipo_MOD_Temp[ii, 6], periodo_Tipo_MOD_Temp[ii, 7], periodo_Tipo_MOD_Temp[ii, 8], periodo_Tipo_MOD_Temp[ii, 9], periodo_Tipo_MOD_Temp[ii, 10], periodo_Tipo_MOD_Temp[ii, 11], periodo_Tipo_MOD_Temp[ii, 12], periodo_Tipo_MOD_Temp[ii, 13], periodo_Tipo_MOD_Temp[ii, 14]);

                // CALCULAR PENETRACION POR HOGAR - SOBRE MODALIDAD
                if (V1 == "Consolidado")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Hogar.Clone();
                }
                else
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Hogar.Clone();
                }
               
                for (int i = 0; i < periodo_Modalidad_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoModalidad == periodo_Modalidad_Temporal[i, 0].ToString())
                    {
                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            periodo_Tipo_MOD_Temp[ii, 2] / periodo_Modalidad_Temporal[i, 1] * 100,
                            periodo_Tipo_MOD_Temp[ii, 3] / periodo_Modalidad_Temporal[i, 2] * 100,
                            periodo_Tipo_MOD_Temp[ii, 4] / periodo_Modalidad_Temporal[i, 3] * 100,
                            periodo_Tipo_MOD_Temp[ii, 5] / periodo_Modalidad_Temporal[i, 4] * 100,
                            periodo_Tipo_MOD_Temp[ii, 6] / periodo_Modalidad_Temporal[i, 5] * 100,
                            periodo_Tipo_MOD_Temp[ii, 7] / periodo_Modalidad_Temporal[i, 6] * 100,
                            periodo_Tipo_MOD_Temp[ii, 8] / periodo_Modalidad_Temporal[i, 7] * 100,
                            periodo_Tipo_MOD_Temp[ii, 9] / periodo_Modalidad_Temporal[i, 8] * 100,
                            periodo_Tipo_MOD_Temp[ii, 10] / periodo_Modalidad_Temporal[i, 9] * 100,
                            periodo_Tipo_MOD_Temp[ii, 11] / periodo_Modalidad_Temporal[i, 10] * 100,
                            periodo_Tipo_MOD_Temp[ii, 12] / periodo_Modalidad_Temporal[i, 11] * 100,
                            periodo_Tipo_MOD_Temp[ii, 13] / periodo_Modalidad_Temporal[i, 12] * 100,
                            periodo_Tipo_MOD_Temp[ii, 14] / periodo_Modalidad_Temporal[i, 13] * 100);
                        break;
                    }
                }

                // CALCULAR PENETRACION POR HOGAR - SOBRE TIPO
                if (V1 == "Consolidado")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Mercado_Tipo_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Lima_Tipo_Hogar.Clone();
                }
                else
                {
                    periodo_Tipo_Temporal = (double[,])periodo_Total_Ciudades_Tipo_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Tipo_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoTipo == periodo_Tipo_Temporal[i, 0].ToString())
                    {
                        Actualizar_BD_Penetracion(V2, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                            periodo_Tipo_MOD_Temp[ii, 2] / periodo_Tipo_Temporal[i, 1] * 100,
                            periodo_Tipo_MOD_Temp[ii, 3] / periodo_Tipo_Temporal[i, 2] * 100,
                            periodo_Tipo_MOD_Temp[ii, 4] / periodo_Tipo_Temporal[i, 3] * 100,
                            periodo_Tipo_MOD_Temp[ii, 5] / periodo_Tipo_Temporal[i, 4] * 100,
                            periodo_Tipo_MOD_Temp[ii, 6] / periodo_Tipo_Temporal[i, 5] * 100,
                            periodo_Tipo_MOD_Temp[ii, 7] / periodo_Tipo_Temporal[i, 6] * 100,
                            periodo_Tipo_MOD_Temp[ii, 8] / periodo_Tipo_Temporal[i, 7] * 100,
                            periodo_Tipo_MOD_Temp[ii, 9] / periodo_Tipo_Temporal[i, 8] * 100,
                            periodo_Tipo_MOD_Temp[ii, 10] / periodo_Tipo_Temporal[i, 9] * 100,
                            periodo_Tipo_MOD_Temp[ii, 11] / periodo_Tipo_Temporal[i, 10] * 100,
                            periodo_Tipo_MOD_Temp[ii, 12] / periodo_Tipo_Temporal[i, 11] * 100,
                            periodo_Tipo_MOD_Temp[ii, 13] / periodo_Tipo_Temporal[i, 12] * 100,
                            periodo_Tipo_MOD_Temp[ii, 14] / periodo_Tipo_Temporal[i, 13] * 100);
                        break;
                    }
                }
            }

            // CODIGO RECORRIDO POR CATEGORIA Y MODALIDAD                              
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_CATEGORIA_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                int rows = 0;
                int columnas = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    columnas = cols;
                    while (reader_1.Read())
                    {
                        for (int i = 0; i < cols; i++) // LEYENDO DESDE LA COLUMNA CON LOS VALORES
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
                                periodo_Total_Mercado_CateMOD_Hogar[rows, i] = valor_1;
                                periodo_Cate_MOD_Temp[rows, i] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_CateMOD_Hogar[rows, i] = valor_1;
                                periodo_Cate_MOD_Temp[rows, i] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_CateMOD_Hogar[rows, i] = valor_1;
                                periodo_Cate_MOD_Temp[rows, i] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            for (int ii = 0; ii < periodo_Cate_MOD_Temp.GetLength(0); ii++) // FILAS
            {
                CodigoCategoria = periodo_Cate_MOD_Temp[ii, 1].ToString();
                for (int i = 0; i < Codigo_Nombres_Categoria.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoCategoria == Codigo_Nombres_Categoria[i, 0])
                    {
                        V1_ = Codigo_Nombres_Categoria[i, 1].Substring(3, Codigo_Nombres_Categoria[i, 1].Length - 3); ;
                        Mercado_ = Codigo_Nombres_Categoria[i, 1];
                        break;
                    }
                }
                CodigoModalidad = periodo_Cate_MOD_Temp[ii, 0].ToString();
                for (int i = 0; i < Codigo_Nombres_Modalidad.Length / 2; i++) // SE DIVIDE ENTRE 2 PQ TIENE 2 DIMENSIONES
                {
                    if (CodigoModalidad == Codigo_Nombres_Modalidad[i, 0])
                    {
                        V2 = Codigo_Nombres_Modalidad[i, 1].Substring(3, Codigo_Nombres_Modalidad[i, 1].Length - 3); ;
                        Mercado = Codigo_Nombres_Modalidad[i, 1];
                        break;
                    }
                }

                Actualizar_BD(V1_, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", periodo_Cate_MOD_Temp[ii, 2], periodo_Cate_MOD_Temp[ii, 3], periodo_Cate_MOD_Temp[ii, 4], periodo_Cate_MOD_Temp[ii, 5], periodo_Cate_MOD_Temp[ii, 6], periodo_Cate_MOD_Temp[ii, 7], periodo_Cate_MOD_Temp[ii, 8], periodo_Cate_MOD_Temp[ii, 9], periodo_Cate_MOD_Temp[ii, 10], periodo_Cate_MOD_Temp[ii, 11], periodo_Cate_MOD_Temp[ii, 12], periodo_Cate_MOD_Temp[ii, 13], periodo_Cate_MOD_Temp[ii, 14]);
                //MODALIDAD CATEGORIA
                Actualizar_BD(V2, "Suma", Variable, Ciudad_, Mercado_, "HOGARES", "MENSUAL", periodo_Cate_MOD_Temp[ii, 2], periodo_Cate_MOD_Temp[ii, 3], periodo_Cate_MOD_Temp[ii, 4], periodo_Cate_MOD_Temp[ii, 5], periodo_Cate_MOD_Temp[ii, 6], periodo_Cate_MOD_Temp[ii, 7], periodo_Cate_MOD_Temp[ii, 8], periodo_Cate_MOD_Temp[ii, 9], periodo_Cate_MOD_Temp[ii, 10], periodo_Cate_MOD_Temp[ii, 11], periodo_Cate_MOD_Temp[ii, 12], periodo_Cate_MOD_Temp[ii, 13], periodo_Cate_MOD_Temp[ii, 14]);

                // CALCULAR PENETRACION POR HOGAR
                if (V1 == "Consolidado")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Mercado_Modalidad_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Lima_Modalidad_Hogar.Clone();
                }
                else
                {
                    periodo_Modalidad_Temporal = (double[,])periodo_Total_Ciudades_Modalidad_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Modalidad_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoModalidad == periodo_Modalidad_Temporal[i, 0].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = periodo_Modalidad_Temporal[i, 1] != 0 ? periodo_Cate_MOD_Temp[ii, 2] / periodo_Modalidad_Temporal[i, 1] * 100 : 0;
                        t2 = periodo_Modalidad_Temporal[i, 2] != 0 ? periodo_Cate_MOD_Temp[ii, 3] / periodo_Modalidad_Temporal[i, 2] * 100 : 0;
                        t3 = periodo_Modalidad_Temporal[i, 3] != 0 ? periodo_Cate_MOD_Temp[ii, 4] / periodo_Modalidad_Temporal[i, 3] * 100 : 0;
                        t4 = periodo_Modalidad_Temporal[i, 4] != 0 ? periodo_Cate_MOD_Temp[ii, 5] / periodo_Modalidad_Temporal[i, 4] * 100 : 0;
                        t5 = periodo_Modalidad_Temporal[i, 5] != 0 ? periodo_Cate_MOD_Temp[ii, 6] / periodo_Modalidad_Temporal[i, 5] * 100 : 0;
                        t6 = periodo_Modalidad_Temporal[i, 6] != 0 ? periodo_Cate_MOD_Temp[ii, 7] / periodo_Modalidad_Temporal[i, 6] * 100 : 0;
                        t7 = periodo_Modalidad_Temporal[i, 7] != 0 ? periodo_Cate_MOD_Temp[ii, 8] / periodo_Modalidad_Temporal[i, 7] * 100 : 0;
                        t8 = periodo_Modalidad_Temporal[i, 8] != 0 ? periodo_Cate_MOD_Temp[ii, 9] / periodo_Modalidad_Temporal[i, 8] * 100 : 0;
                        t9 = periodo_Modalidad_Temporal[i, 9] != 0 ? periodo_Cate_MOD_Temp[ii, 10] / periodo_Modalidad_Temporal[i, 9] * 100 : 0;
                        t10 = periodo_Modalidad_Temporal[i, 10] != 0 ? periodo_Cate_MOD_Temp[ii, 11] / periodo_Modalidad_Temporal[i, 10] * 100 : 0;
                        t11 = periodo_Modalidad_Temporal[i, 11] != 0 ? periodo_Cate_MOD_Temp[ii, 12] / periodo_Modalidad_Temporal[i, 11] * 100 : 0;
                        t12 = periodo_Modalidad_Temporal[i, 12] != 0 ? periodo_Cate_MOD_Temp[ii, 13] / periodo_Modalidad_Temporal[i, 12] * 100 : 0;
                        t13 = periodo_Modalidad_Temporal[i, 13] != 0 ? periodo_Cate_MOD_Temp[ii, 14] / periodo_Modalidad_Temporal[i, 13] * 100 : 0;

                        Actualizar_BD_Penetracion(V1_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        
                        break;

                    }
                }
                // CALCULAR PENETRACION POR HOGAR - CATEGORIA POR MODALIDAD
                if (V1 == "Consolidado")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Mercado_Categoria_Hogar.Clone();
                }
                else if (V1 == "Lima")
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Lima_Categoria_Hogar.Clone();
                }
                else
                {
                    periodo_Cat_Temporal = (double[,])periodo_Total_Ciudades_Categoria_Hogar.Clone();
                }

                for (int i = 0; i < periodo_Cat_Temporal.GetLength(0); i++) // GetLength(0) lee numero de filas array
                {
                    if (CodigoCategoria == periodo_Cat_Temporal[i, 0].ToString())
                    {
                        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
                        t1 = periodo_Cat_Temporal[i, 1] != 0 ? periodo_Cate_MOD_Temp[ii, 2] / periodo_Cat_Temporal[i, 1] * 100 : 0;
                        t2 = periodo_Cat_Temporal[i, 2] != 0 ? periodo_Cate_MOD_Temp[ii, 3] / periodo_Cat_Temporal[i, 2] * 100 : 0;
                        t3 = periodo_Cat_Temporal[i, 3] != 0 ? periodo_Cate_MOD_Temp[ii, 4] / periodo_Cat_Temporal[i, 3] * 100 : 0;
                        t4 = periodo_Cat_Temporal[i, 4] != 0 ? periodo_Cate_MOD_Temp[ii, 5] / periodo_Cat_Temporal[i, 4] * 100 : 0;
                        t5 = periodo_Cat_Temporal[i, 5] != 0 ? periodo_Cate_MOD_Temp[ii, 6] / periodo_Cat_Temporal[i, 5] * 100 : 0;
                        t6 = periodo_Cat_Temporal[i, 6] != 0 ? periodo_Cate_MOD_Temp[ii, 7] / periodo_Cat_Temporal[i, 6] * 100 : 0;
                        t7 = periodo_Cat_Temporal[i, 7] != 0 ? periodo_Cate_MOD_Temp[ii, 8] / periodo_Cat_Temporal[i, 7] * 100 : 0;
                        t8 = periodo_Cat_Temporal[i, 8] != 0 ? periodo_Cate_MOD_Temp[ii, 9] / periodo_Cat_Temporal[i, 8] * 100 : 0;
                        t9 = periodo_Cat_Temporal[i, 9] != 0 ? periodo_Cate_MOD_Temp[ii, 10] / periodo_Cat_Temporal[i, 9] * 100 : 0;
                        t10 = periodo_Cat_Temporal[i, 10] != 0 ? periodo_Cate_MOD_Temp[ii, 11] / periodo_Cat_Temporal[i, 10] * 100 : 0;
                        t11 = periodo_Cat_Temporal[i, 11] != 0 ? periodo_Cate_MOD_Temp[ii, 12] / periodo_Cat_Temporal[i, 11] * 100 : 0;
                        t12 = periodo_Cat_Temporal[i, 12] != 0 ? periodo_Cate_MOD_Temp[ii, 13] / periodo_Cat_Temporal[i, 12] * 100 : 0;
                        t13 = periodo_Cat_Temporal[i, 13] != 0 ? periodo_Cate_MOD_Temp[ii, 14] / periodo_Cat_Temporal[i, 13] * 100 : 0;

                        Actualizar_BD_Penetracion(V2, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                        break;
                    }
                }
            }
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
        private void Actualizar_BD_Penetracion(string _V1, string _V2, string _Variable, string _Ciudad, string _Mercado, string _Unidad, string _Reporte, double _ANO_0, double _ANO_1, double _ANO_2, double _PER_12M_1, double _PER_12M_2, double _PER_6M_1, double _PER_6M_2, double _PER_3M_1, double _PER_3M_2, double _PER_1M_1, double _PER_1M_2, double _PER_YTD_1, double _PER_YTD_2)
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


                double Dato_1 = 0, Dato_2 = 0, Dato_3 = 0, Dato_4 = 0, Dato_5 = 0, Dato_6 = 0;
                if (_ANO_1 != 0)
                {
                    Dato_1 = (_ANO_2 - _ANO_1 ) ;
                }
                if (_PER_12M_1 != 0)
                {
                    Dato_2 = (_PER_12M_2 - _PER_12M_1 ) ;
                }
                if (_PER_6M_1 != 0)
                {
                    Dato_3 = (_PER_6M_2 - _PER_6M_1 );
                }
                if (_PER_3M_1 != 0)
                {
                    Dato_4 = (_PER_3M_2 - _PER_3M_1 ) ;
                }
                if (_PER_1M_1 != 0)
                {
                    Dato_5 = (_PER_1M_2 - _PER_1M_1 ) ;
                }
                if (_PER_YTD_1 != 0)
                {
                    Dato_6 = (_PER_YTD_2 - _PER_YTD_1 ) ;
                }

                db_Zoho.AddInParameter(cmd_1, "_VAR_ANO", DbType.Double, Dato_1);
                db_Zoho.AddInParameter(cmd_1, "_VAR_12M", DbType.Double, Dato_2);
                db_Zoho.AddInParameter(cmd_1, "_VAR_6M", DbType.Double, Dato_3);
                db_Zoho.AddInParameter(cmd_1, "_VAR_3M", DbType.Double, Dato_4);
                db_Zoho.AddInParameter(cmd_1, "_VAR_1M", DbType.Double, Dato_5);
                db_Zoho.AddInParameter(cmd_1, "_VAR_YTD", DbType.Double, Dato_6);

                db_Zoho.ExecuteNonQuery(cmd_1);
            }
        }
    }
}