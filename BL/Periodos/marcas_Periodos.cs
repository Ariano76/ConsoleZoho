﻿using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.CodeDom;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Policy;

namespace BL.Periodos
{
    public class Marcas_Periodos
    {
        private int[] Codigo_MARCA_VD = { 540, 5914, 1163, 504, 24, 1764, 318, 3206, 8019, 8420 }; // MARCAS PREDEFINIDAS VD 
        public string[] Codigo_MARCA_VR = new string[5];
        public string[] Codigo_MARCA_Nombre_VR = new string[5];

        private int[] Codigo_Belcorp = { 540, 1163, 5914 }; // GRUPO BELCORP
        private int[] Codigo_Loreal = { 5193, 696, 63, 1415, 3983, 1620, 3132, 176, 189, 1102, 2968, 6595, 1271, 293, 279, 282, 322, 583, 5434, 424, 423, 4225, 506, 1350, 8493, 513, 537, 209 }; // GRUPO LOREAL
        private int[] Codigo_Lauder = { 4825,10391,6881,8698,81,6305,4779,6197,2605,8535,4200,284,4522,10647,8009 }; // GRUPO ESTEE LAUDER
            
        private int[] Codigos_NSE = { 488, 489, 490, 491, 492 }; // NSE
        private int[] Codigo_Tipos = { 229, 238, 164, 158, 161, 174, 215, 178, 202, 237, 226, 228, 173, 235, 175 }; // TIPOS 
        int[] Cod_NSE_Temp = new int[5]; // NSE
        int[,] Cod_Tipo_Temp = new int[75, 2];

        int[,] Codigo_Tipos_Categorias = new int[15, 2] { { 229, 127 }, { 238, 127 }, { 164, 123 }, { 158, 123 }, { 161, 123 }, { 175, 124 }, { 174, 124 }, { 215, 126 }, { 178, 124 }, { 202, 125 }, { 237, 127 }, { 226, 127 }, { 228, 127 }, { 173, 124 }, { 235, 127 } };  // codigos de tipos y categorias
        string[,] Codigo_Nombres_Tipo = new string[15, 2] { { "229", "12. Acondic.Adultos-Niños-Bb" }, { "238", "15. Aerosol" }, { "164", "03. Colonia Baño" }, { "158", "01. Colonia Femeninas" }, { "161", "02. Colonia Masculinas" }, { "175", "06. Delineador Ojos" }, { "174", "05. Embellecedor-Rimmel" }, { "215", "09. Humectante/Nutritiva Corporal" }, { "178", "07. Labiales" }, { "202", "08. Nutritiva Revit. Facial" }, { "237", "14. Roll-On" }, { "226", "10. Shampoo Adultos" }, { "228", "11. Shampoo Bebes" }, { "173", "04. Sombras" }, { "235", "13. Trat.Capilar Adultos/Niños" } };  // codigos y nombre de tipos
        string[,] Codigo_Nombres_NSE = new string[5, 2] { { "488", "Alto" }, { "489", "Medio" }, { "490", "Medio Bajo" }, { "491", "Bajo" }, { "492", "Muy Bajo" } };  // Codigos y Nombre de NSE
        string[,] Codigo_Nombres_Categoria = new string[5, 2] { { "123", "1. Fragancias" }, { "124", "2. Maquillaje" }, { "125", "3. Tratamiento Facial" }, { "126", "4. Tratamiento Corporal" }, { "127", "5. Cuidado Personal" } };  // Codigos y Nombre de Categoria
        string[,] Codigo_Nombres_Modalidad = new string[2, 2] { { "87", "1. VD" }, { "88", "2. VR" } }; // Codigos y Nombre de Modalidad
        private int[] Codigo_Categoria_All = { 123, 124, 125, 126, 127 }; // Categorias   
        private int[] Codigo_Categoria = { 123, 124, 127 }; // Categorias que se miden mensualmente
        private int[] Codigo_Modalidad_Venta = { 87, 88 }; // Modalidad de Venta 


        public double[] sShareValor;

        public string Codigo_Marcas_3M_Tipo, Codigo_Marcas_3M_Categoria, Codigo_Marcas_3M_Total_Cosmeticos;
        public string Codigo_Grupo_Belcorp_3M_Tipo, Codigo_Grupo_Loreal_3M_Tipo, Codigo_Grupo_Lauder_3M_Tipo;
        public string Codigo_cadena_NSE;
        public string Codigo_cadena_Categoria;
        public string Codigo_cadena_Tipos, Codigo_cadena_Modalidad;


        public double[,] periodo_Total_Mercado_Valor = new double[1, 13];  // TOTAL VALORES
        public double[,] periodo_Total_Lima_Valor = new double[1, 13];  // LIMA  VALORES
        public double[,] periodo_Total_Ciudades_Valor = new double[1, 13];  // CIUDAD VALORES
        public double[,] periodo_Total_Mercado_Categoria_Valor = new double[5, 14];  // CATEG TOTAL NSE Valor
        public double[,] periodo_Total_Lima_Categoria_Valor = new double[5, 14];  // CATEG LIMA NSE Valor
        public double[,] periodo_Total_Ciudades_Categoria_Valor = new double[5, 14];  // CATEG CIUDADES NSE Valor
        public double[,] periodo_Total_Mercado_Tipo_Valor = new double[15, 14];  // TIPO TOTAL NSE Valor
        public double[,] periodo_Total_Lima_Tipo_Valor = new double[15, 14];  // TIPO LIMA NSE Valor
        public double[,] periodo_Total_Ciudades_Tipo_Valor = new double[15, 14];  // TIPO CIUDADES NSE Valor

        public double[,] periodo_Total_Mercado_UU = new double[1, 13];  // TOTAL UU
        public double[,] periodo_Total_Lima_UU = new double[1, 13];  // LIMA  UU
        public double[,] periodo_Total_Ciudades_UU = new double[1, 13];  // CIUDAD UU
        public double[,] periodo_Total_Mercado_Categoria_UU = new double[5, 14];  // CATEG TOTAL NSE UU
        public double[,] periodo_Total_Lima_Categoria_UU = new double[5, 14];  // CATEG LIMA NSE UU
        public double[,] periodo_Total_Ciudades_Categoria_UU = new double[5, 14];  // CATEG CIUDADES NSE UU
        public double[,] periodo_Total_Mercado_Tipo_UU = new double[15, 14];  // TIPO TOTAL NSE UU
        public double[,] periodo_Total_Lima_Tipo_UU = new double[15, 14];  // TIPO LIMA NSE UU
        public double[,] periodo_Total_Ciudades_Tipo_UU = new double[15, 14];  // TIPO CIUDADES NSE UU

        public double[,] Total_Mercado_Marca_Valor = new double[15, 14];  // TOTAL MARCA VALOR
        public double[,] Total_Lima_Marca_Valor = new double[15, 14];  // LIMA MARCA VALOR
        public double[,] Total_Ciudades_Marca_Valor = new double[15, 14];  // CIUDADES MARCA VALOR
        public double[,] Categ_Mercado_Marca_Valor = new double[15, 15];  // CATEG TOTAL MARCA VALOR
        public double[,] Categ_Lima_Marca_Valor = new double[15, 15];  // CATEG LIMA MARCA VALOR
        public double[,] Categ_Ciudades_Marca_Valor = new double[15, 15];  // CATEG CIUDADES MARCA VALOR
        public double[,] Tipo_Mercado_Marca_Valor = new double[15, 15];  // TIPO TOTAL MARCA VALOR
        public double[,] Tipo_Lima_Marca_Valor = new double[15, 15];  // TIPO LIMA MARCA VALOR
        public double[,] Tipo_Ciudades_Marca_Valor = new double[15, 15];  // TIPO CIUDADES MARCA VALOR

        public double[,] Total_Mercado_Marca_UU = new double[15, 14];  // TOTAL MARCA UU
        public double[,] Total_Lima_Marca_UU = new double[15, 14];  // LIMA MARCA UU
        public double[,] Total_Ciudades_Marca_UU = new double[15, 14];  // CIUDADES MARCA UU
        public double[,] Categ_Mercado_Marca_UU = new double[15, 15];  // CATEG TOTAL MARCA UU
        public double[,] Categ_Lima_Marca_UU = new double[15, 15];  // CATEG LIMA MARCA UU
        public double[,] Categ_Ciudades_Marca_UU = new double[15, 15];  // CATEG CIUDADES MARCA UU
        public double[,] Tipo_Mercado_Marca_UU = new double[15, 15];  // TIPO TOTAL MARCA UU
        public double[,] Tipo_Lima_Marca_UU = new double[15, 15];  // TIPO LIMA MARCA UU
        public double[,] Tipo_Ciudades_Marca_UU = new double[15, 15];  // TIPO CIUDADES MARCA UU


        public double[,] Temp_Total_Valor = new double[1, 13];  // TEMPORAL TOTAL VALOR
        public double[,] Temp_Categoria_Valor = new double[5, 14];  // CATEGORIA VALOR
        public double[,] Temp_Tipo_Valor = new double[15, 14];  // TIPO VALOR
        public double[,] Temp_Total_Marca_Valor = new double[15, 14];  // TOTAL MARCA VALOR
        public double[,] Temp_Categ_Marca_Valor = new double[15, 15];  // CATEGORIA MARCA VALOR
        public double[,] Temp_Tipo_Marca_Valor = new double[15, 15];  // TIPO MARCA VALOR

        public double[,] Temp_Total_UU = new double[1, 13];  // TEMPORAL TOTAL UU
        public double[,] Temp_Categoria_UU = new double[5, 14];  // CATEGORIA UU
        public double[,] Temp_Tipo_UU = new double[15, 14];  // TIPO UU
        public double[,] Temp_Total_Marca_UU = new double[15, 14];  // TOTAL MARCA UU
        public double[,] Temp_Categ_Marca_UU = new double[15, 15];  // CATEGORIA MARCA UU
        public double[,] Temp_Tipo_Marca_UU = new double[15, 15];  // TIPO MARCA UU

        public double[,] Total_Mcdo_Hogar_Pais = new double[1, 13];  // TOTAL HOGAR PAIS
        public double[,] Total_Mcdo_Hogar_Lima = new double[1, 13];  // TOTAL HOGAR PAIS
        public double[,] Total_Mcdo_Hogar_Ciudad = new double[1, 13];  // TOTAL HOGAR PAIS
        public double[,] Total_Mcdo_Hogar_Temp = new double[1, 13];  // TOTAL HOGAR PAIS

        //public double[,] periodo_Temp_Unidad = new double[1, 13];  // TEMPORAL POR  UNIDAD              

        public double[,] periodo_Universos_NSE = new double[15, 15];  // UNIVERSOS        
        public double[,] periodo_Universos = new double[3, 15];  // UNIVERSOS       



        private readonly DateTime[] Periodos = new DateTime[7];
        string V1, V1_, V1__, V1_M, V2, V2_, Ciudad_, Mercado, Mercado_, Variable, VariableGM, Variable_Promedio, Unidad;
        public string resultadoBD;
        double valor_1;
        double t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13;
        double r1, r2, r3, r4, r5, r6, r7, r8, r9, r10, r11, r12, r13;
        public static DateTime HoraStart;
        TimeSpan HoraFin;

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
        public void Recuperar_Marcas_Grupo_Belcorp()
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Belcorp.GetLength(0); i++)
            {
                xIdMarca += Codigo_Belcorp[i].ToString() + ",";
            }
            Codigo_Grupo_Belcorp_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }
        public void Recuperar_Marcas_Grupo_Loreal()
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Loreal.GetLength(0); i++)
            {
                xIdMarca += Codigo_Loreal[i].ToString() + ",";
            }
            Codigo_Grupo_Loreal_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }
        public void Recuperar_Marcas_Grupo_Lauder()
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_Lauder.GetLength(0); i++)
            {
                xIdMarca += Codigo_Lauder[i].ToString() + ",";
            }
            Codigo_Grupo_Lauder_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
        }

        public void Recuperar_Marcas_Top_5_Retail_x_Tipo(string xCiudad, int xTipo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Tipo = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_TIPO_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Categoria(string xCiudad, int xCateg, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Categoria = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_CATEG_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCateg);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(string xCiudad, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_TOTAL_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Tipo_UU(string xCiudad, int xTipo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Tipo = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_TIPO_3M_UU"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Categoria_UU(string xCiudad, int xCateg, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Categoria = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_CATEG_3M_UU"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCateg);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_UU(string xCiudad, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_TOTAL_3M_UU"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Hogar_Total(string xCiudad, string xCategoria)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOP_5_TOTAL_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCAS", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCategoria);
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
        public void Recuperar_Marcas_Top_5_Retail_x_Hogar_Total_Tipo(string xCiudad, string xTipo)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOP_5_TIPO_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCAS", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);
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
        public void Recuperar_Marcas_Top_5_Retail_x_GastoMedio_Total(string xCiudad, string xPeriodo, string xCategoria, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_GM_VALOR_TOP_5_TOTAL_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodo);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));                
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCategoria);
                db_Zoho.AddInParameter(cmd_1, "_IDMONEDA", DbType.Int16, xMoneda);
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
        public void Recuperar_Marcas_Top_5_x_PPU_Total_Cosmeticos(string xCiudad, String xPeriodo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            Codigo_Marcas_3M_Total_Cosmeticos = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_PPU_TOTAL_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, xPeriodo);
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
        public void Recuperar_Marcas_Top_5_x_PPU_Categoria(string xCiudad, int xCateg, String xPeriodo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Categoria = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_PPU_CATEG_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCateg);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, xPeriodo);
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
        public void Recuperar_Marcas_Top_5_x_PPU_Tipo(string xCiudad, int xTipo, String xPeriodo, int xMoneda)
        {
            string xIdMarca = "";
            for (int i = 0; i < Codigo_MARCA_VD.GetLength(0); i++)
            {
                xIdMarca += Codigo_MARCA_VD[i].ToString() + ",";
            }
            //Codigo_Marcas_3M_Tipo = xIdMarca.Substring(0, xIdMarca.Length - 1);
            Codigo_Marcas_3M_Tipo = xIdMarca.ToString();

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOP_5_PPU_TIPO_3M"))
            {
                int cont = 0;
                string xIdMarcaOtros = "";
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);
                db_Zoho.AddInParameter(cmd_1, "_IDMARCA", DbType.String, xIdMarca.Substring(0, xIdMarca.Length - 1));
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_PER3M_2", DbType.String, xPeriodo);
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

        // VALORES MONEDA LOCAL - DOLARES
        public void Periodos_Cosmeticos_Total_Valores(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, [Optional] int NumMeses)
        {
            Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xCiudad, 1);

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
            
            HoraStart = DateTime.Now;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_MERCADO"))
            {
                //db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
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
                                Temp_Total_Valor[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Valor[rows, i - 1] = valor_1;
                                Temp_Total_Valor[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Valor[rows, i - 1] = valor_1;
                                Temp_Total_Valor[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
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
                //int rows = 0;
                
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {        
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Total_Valor[0, 0]*100);
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Total_Valor[0, 1]*100);
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Total_Valor[0, 2]*100);
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Total_Valor[0, 3]*100);
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Total_Valor[0, 4]*100);
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Total_Valor[0, 5]*100);
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Total_Valor[0, 6]*100);
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Total_Valor[0, 7]*100);
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Total_Valor[0, 8]*100);
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Total_Valor[0, 9]*100);
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Total_Valor[0, 10]*100);
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Total_Valor[0, 11]*100);
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Total_Valor[0, 12]*100);

                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD_Penetracion(V1_M, "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
         
            // BELCORP
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_Valor[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_Valor[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_Valor[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_Valor[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_Valor[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_Valor[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_Valor[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_Valor[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_Valor[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_Valor[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_Valor[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_Valor[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_Valor[0, 12] * 100;

                        Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
            // LOREAL
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_Valor[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_Valor[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_Valor[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_Valor[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_Valor[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_Valor[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_Valor[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_Valor[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_Valor[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_Valor[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_Valor[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_Valor[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_Valor[0, 12] * 100;

                        Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
            // ESTEE LAUDER
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_Valor[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_Valor[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_Valor[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_Valor[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_Valor[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_Valor[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_Valor[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_Valor[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_Valor[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_Valor[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_Valor[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_Valor[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_Valor[0, 12] * 100;

                        Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE DOLARES", Ciudad_, "0. Cosmeticos", "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
            
            TimeSpan HoraFin;
            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Mercado : " + HoraFin.ToString());

          
            #region CATEGORIA
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
                                periodo_Total_Mercado_Categoria_Valor[rows, i - 1] = valor_1;
                                Temp_Categoria_Valor[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_Valor[rows, i - 1] = valor_1;
                                Temp_Categoria_Valor[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_Valor[rows, i - 1] = valor_1;
                                Temp_Categoria_Valor[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }

            int posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Categoria.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_Retail_x_Categoria(xCiudad, int.Parse(Codigo_Nombres_Categoria[x, 0].ToString()), 1);

                for (int i = 0; i < Temp_Categoria_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Categoria[x, 0] == Temp_Categoria_Valor[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Categoria[x, 1];
                        break;
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Categoria_Valor[posicion, 1]*100);
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Categoria_Valor[posicion, 2]*100);
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Categoria_Valor[posicion, 3]*100);
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Categoria_Valor[posicion, 4]*100);
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Categoria_Valor[posicion, 5]*100);
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Categoria_Valor[posicion, 6]*100);
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Categoria_Valor[posicion, 7]*100);
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Categoria_Valor[posicion, 8]*100);
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Categoria_Valor[posicion, 9]*100);
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Categoria_Valor[posicion, 10]*100);
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Categoria_Valor[posicion, 11]*100);
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Categoria_Valor[posicion, 12]*100);
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()) / Temp_Categoria_Valor[posicion, 13]*100);

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD_Penetracion(V1_M, "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                            rows++;
                        }                      
                    }
                }               

                // BELCORP
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = double.Parse(reader_1[1].ToString()) / Temp_Categoria_Valor[posicion, 1] * 100;
                            t2 = double.Parse(reader_1[2].ToString()) / Temp_Categoria_Valor[posicion, 2] * 100;
                            t3 = double.Parse(reader_1[3].ToString()) / Temp_Categoria_Valor[posicion, 3] * 100;
                            t4 = double.Parse(reader_1[4].ToString()) / Temp_Categoria_Valor[posicion, 4] * 100;
                            t5 = double.Parse(reader_1[5].ToString()) / Temp_Categoria_Valor[posicion, 5] * 100;
                            t6 = double.Parse(reader_1[6].ToString()) / Temp_Categoria_Valor[posicion, 6] * 100;
                            t7 = double.Parse(reader_1[7].ToString()) / Temp_Categoria_Valor[posicion, 7] * 100;
                            t8 = double.Parse(reader_1[8].ToString()) / Temp_Categoria_Valor[posicion, 8] * 100;
                            t9 = double.Parse(reader_1[9].ToString()) / Temp_Categoria_Valor[posicion, 9] * 100;
                            t10 = double.Parse(reader_1[10].ToString()) / Temp_Categoria_Valor[posicion, 10] * 100;
                            t11 = double.Parse(reader_1[11].ToString()) / Temp_Categoria_Valor[posicion, 11] * 100;
                            t12 = double.Parse(reader_1[12].ToString()) / Temp_Categoria_Valor[posicion, 12] * 100;
                            t13 = double.Parse(reader_1[13].ToString()) / Temp_Categoria_Valor[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // LOREAL
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Categoria_Valor[posicion, 1] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Categoria_Valor[posicion, 2] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Categoria_Valor[posicion, 3] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Categoria_Valor[posicion, 4] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Categoria_Valor[posicion, 5] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Categoria_Valor[posicion, 6] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Categoria_Valor[posicion, 7] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Categoria_Valor[posicion, 8] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Categoria_Valor[posicion, 9] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Categoria_Valor[posicion, 10] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Categoria_Valor[posicion, 11] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Categoria_Valor[posicion, 12] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Categoria_Valor[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // ESTEE LAUDER
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Categoria_Valor[0, 0] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Categoria_Valor[0, 1] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Categoria_Valor[0, 2] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Categoria_Valor[0, 3] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Categoria_Valor[0, 4] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Categoria_Valor[0, 5] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Categoria_Valor[0, 6] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Categoria_Valor[0, 7] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Categoria_Valor[0, 8] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Categoria_Valor[0, 9] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Categoria_Valor[0, 10] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Categoria_Valor[0, 11] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Categoria_Valor[0, 12] * 100;

                            Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }

            }
            
            //TimeSpan HoraFin;
            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Categoria : " + HoraFin.ToString());

            #endregion CATEGORIA

            #region TIPO
            HoraStart = DateTime.Now;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_TOTAL_TIPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
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
                                periodo_Total_Mercado_Tipo_Valor[rows, i - 1] = valor_1;
                                Temp_Tipo_Valor[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Tipo_Valor[rows, i - 1] = valor_1;
                                Temp_Tipo_Valor[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Tipo_Valor[rows, i - 1] = valor_1;
                                Temp_Tipo_Valor[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }
            
            posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Tipo.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_Retail_x_Tipo(xCiudad, int.Parse(Codigo_Nombres_Tipo[x, 0].ToString()), 1);

                for (int i = 0; i < Temp_Tipo_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Tipo[x, 0] == Temp_Tipo_Valor[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Tipo[x, 1];
                        break;
                    }
                }
                double[,] Temp_Tipo_Marca_Valor = new double[15, 15];  // TIPO MARCA VALORES
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);                    
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                    
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
                    
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        
                        while (reader_1.Read())
                        {                            
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Tipo_Valor[posicion, 1] * 100);
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Tipo_Valor[posicion, 2] * 100);
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Tipo_Valor[posicion, 3] * 100);
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Tipo_Valor[posicion, 4] * 100);
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Tipo_Valor[posicion, 5] * 100);
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Tipo_Valor[posicion, 6] * 100);
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Tipo_Valor[posicion, 7] * 100);
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Tipo_Valor[posicion, 8] * 100);
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Tipo_Valor[posicion, 9] * 100);
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Tipo_Valor[posicion, 10] * 100);
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Tipo_Valor[posicion, 11] * 100);
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Tipo_Valor[posicion, 12] * 100);
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()) / Temp_Tipo_Valor[posicion, 13] * 100);

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD_Penetracion(V1_M, "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }                
                // BELCORP
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                    
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = double.Parse(reader_1[1].ToString()) / Temp_Tipo_Valor[posicion, 1] * 100;
                            t2 = double.Parse(reader_1[2].ToString()) / Temp_Tipo_Valor[posicion, 2] * 100;
                            t3 = double.Parse(reader_1[3].ToString()) / Temp_Tipo_Valor[posicion, 3] * 100;
                            t4 = double.Parse(reader_1[4].ToString()) / Temp_Tipo_Valor[posicion, 4] * 100;
                            t5 = double.Parse(reader_1[5].ToString()) / Temp_Tipo_Valor[posicion, 5] * 100;
                            t6 = double.Parse(reader_1[6].ToString()) / Temp_Tipo_Valor[posicion, 6] * 100;
                            t7 = double.Parse(reader_1[7].ToString()) / Temp_Tipo_Valor[posicion, 7] * 100;
                            t8 = double.Parse(reader_1[8].ToString()) / Temp_Tipo_Valor[posicion, 8] * 100;
                            t9 = double.Parse(reader_1[9].ToString()) / Temp_Tipo_Valor[posicion, 9] * 100;
                            t10 = double.Parse(reader_1[10].ToString()) / Temp_Tipo_Valor[posicion, 10] * 100;
                            t11 = double.Parse(reader_1[11].ToString()) / Temp_Tipo_Valor[posicion, 11] * 100;
                            t12 = double.Parse(reader_1[12].ToString()) / Temp_Tipo_Valor[posicion, 12] * 100;
                            t13 = double.Parse(reader_1[13].ToString()) / Temp_Tipo_Valor[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // LOREAL
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Tipo_Valor[posicion, 1] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Tipo_Valor[posicion, 2] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Tipo_Valor[posicion, 3] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Tipo_Valor[posicion, 4] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Tipo_Valor[posicion, 5] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Tipo_Valor[posicion, 6] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Tipo_Valor[posicion, 7] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Tipo_Valor[posicion, 8] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Tipo_Valor[posicion, 9] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Tipo_Valor[posicion, 10] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Tipo_Valor[posicion, 11] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Tipo_Valor[posicion, 12] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Tipo_Valor[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // ESTEE LAUDER
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Tipo_Valor[0, 0] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Tipo_Valor[0, 1] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Tipo_Valor[0, 2] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Tipo_Valor[0, 3] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Tipo_Valor[0, 4] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Tipo_Valor[0, 5] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Tipo_Valor[0, 6] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Tipo_Valor[0, 7] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Tipo_Valor[0, 8] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Tipo_Valor[0, 9] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Tipo_Valor[0, 10] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Tipo_Valor[0, 11] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Tipo_Valor[0, 12] * 100;

                            Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE DOLARES", Ciudad_, Mercado, "DOLARES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
            }
            
            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Tipo : " + HoraFin.ToString());

            #endregion TIPO

            #region PPU
            HoraStart = DateTime.Now;

            //PPU MONEDA LOCAL - TOTAL
            Recuperar_Marcas_Top_5_x_PPU_Total_Cosmeticos(xCiudad, _PER3M_2, 1);
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
                //int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()));
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));

                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD(V1_M, "Suma", Unidad, Ciudad_, "0. Cosmeticos", Unidad, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        //rows++;
                    }
                }
            }

            //PPU MONEDA DOLARES - TOTAL
            Recuperar_Marcas_Top_5_x_PPU_Total_Cosmeticos(xCiudad, _PER3M_2, 2);
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, 2);
                //int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()));
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));

                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD(V1_M, "Suma", "PPU (DOL.)", Ciudad_, "0. Cosmeticos", "PPU (DOL.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }                                  
            Grupos(xCiudad, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
            Grupos(xCiudad, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
            Grupos(xCiudad, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
            Grupos(xCiudad, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
            Grupos(xCiudad, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");
            Grupos(xCiudad, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");

            //PPU MONEDA LOCAL - CATEGORIA           
            for (int x = 0; x < Codigo_Nombres_Categoria.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_x_PPU_Categoria(xCiudad, int.Parse(Codigo_Nombres_Categoria[x, 0].ToString()), _PER3M_2, 1);

                for (int i = 0; i < Temp_Categoria_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Categoria[x, 0] == Temp_Categoria_Valor[i, 0].ToString())
                    {
                        //posicion = i;
                        Mercado = Codigo_Nombres_Categoria[x, 1];
                        break;
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                    
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()));

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD(V1_M, "Suma", "PPU (ML)", Ciudad_, Mercado, "PPU (ML)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);                            
                        }
                    }
                }

                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");
                Grupos_Categoria(xCiudad, Codigo_Nombres_Categoria[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");

            }

            //PPU MONEDA DOLARES - CATEGORIA
            for (int x = 0; x < Codigo_Nombres_Categoria.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_x_PPU_Categoria(xCiudad, int.Parse(Codigo_Nombres_Categoria[x, 0].ToString()), _PER3M_2, 2);

                for (int i = 0; i < Temp_Categoria_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Categoria[x, 0] == Temp_Categoria_Valor[i, 0].ToString())
                    {
                        //posicion = i;
                        Mercado = Codigo_Nombres_Categoria[x, 1];
                        break;
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, 2);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()));

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD(V1_M, "Suma", "PPU (DOL.)", Ciudad_, Mercado, "PPU (DOL.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
            }

            //PPU MONEDA LOCAL - TIPO   
            posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Tipo.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_x_PPU_Tipo(xCiudad, int.Parse(Codigo_Nombres_Tipo[x, 0].ToString()), _PER3M_2, 1);

                for (int i = 0; i < Temp_Tipo_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Tipo[x, 0] == Temp_Tipo_Valor[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Tipo[x, 1];
                        break;
                    }
                }
                double[,] Temp_Tipo_Marca_Valor = new double[15, 15];  // TIPO MARCA VALORES
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
                    int rows = 0;

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;

                        while (reader_1.Read())
                        {
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()));

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD(V1_M, "Suma", "PPU (ML)", Ciudad_, Mercado, "PPU (ML)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }

                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Belcorp_3M_Tipo, "00.Belcorp");
                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Loreal_3M_Tipo, "50.Total L'Oreal");
                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 1, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");
                Grupos_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0], Mercado, xCab, xAños, 2, _PER12M_1, _PER12M_2, _PER6M_1, _PER6M_2, _PER3M_1, _PER3M_2, _PER1M_1, _PER3M_2, _PERYTDM_1, _PERYTDM_2, Codigo_Grupo_Lauder_3M_Tipo, "51.Total Estee Lauder");
            }

            //PPU MONEDA DOLARES - TIPO   
            posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Tipo.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_x_PPU_Tipo(xCiudad, int.Parse(Codigo_Nombres_Tipo[x, 0].ToString()), _PER3M_2, 2);

                for (int i = 0; i < Temp_Tipo_Valor.Length; i++)
                {
                    if (Codigo_Nombres_Tipo[x, 0] == Temp_Tipo_Valor[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Tipo[x, 1];
                        break;
                    }
                }
                double[,] Temp_Tipo_Marca_Valor = new double[15, 15];  // TIPO MARCA VALORES
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, 2);
                    int rows = 0;

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;

                        while (reader_1.Read())
                        {                         
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString()));

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD(V1_M, "Suma", "PPU (DOL.)", Ciudad_, Mercado, "PPU (DOL.)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
            }
            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total PPU : " + HoraFin.ToString());

            #endregion
   
        }

        public void Grupos(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string xMarcas_Grupo, string xTituloMarca)
        {
            if (xMoneda == 1) {
                Unidad = "PPU (ML)"; }
            else            {
                Unidad = "PPU (DOL.)"; }

            if (xCiudad == "1") {
                Ciudad_ = "1. Capital"; }
            else if (xCiudad == "1,2,5") {
                Ciudad_ = "0. Consolidado"; }
            else { Ciudad_ = "2. Ciudades"; }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_MERCADO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, xMarcas_Grupo);
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

                            switch (i)
                            {
                                case 1: t1 = valor_1; break;
                                case 2: t2 = valor_1; break;
                                case 3: t3 = valor_1; break;
                                case 4: t4 = valor_1; break;
                                case 5: t5 = valor_1; break;
                                case 6: t6 = valor_1; break;
                                case 7: t7 = valor_1; break;
                                case 8: t8 = valor_1; break;
                                case 9: t9 = valor_1; break;
                                case 10: t10 = valor_1; break;
                                case 11: t11 = valor_1; break;
                                case 12: t12 = valor_1; break;
                                case 13: t13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(xTituloMarca, "Suma", Unidad, Ciudad_, "0. Cosmeticos", Unidad, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
        }

        public void Grupos_Categoria(string xCiudad, string xCodCategoria, string xMercado, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string xMarcas_Grupo, string xTituloMarca)
        {
            if (xMoneda == 1)
            {
                Unidad = "PPU (ML)";
            }
            else
            {
                Unidad = "PPU (DOL.)";
            }

            if (xCiudad == "1")
            {
                Ciudad_ = "1. Capital";
            }
            else if (xCiudad == "1,2,5")
            {
                Ciudad_ = "0. Consolidado";
            }
            else { Ciudad_ = "2. Ciudades"; }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_CATEGORIA_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, xMarcas_Grupo);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCodCategoria);
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

                            switch (i)
                            {
                                case 1: t1 = valor_1; break;
                                case 2: t2 = valor_1; break;
                                case 3: t3 = valor_1; break;
                                case 4: t4 = valor_1; break;
                                case 5: t5 = valor_1; break;
                                case 6: t6 = valor_1; break;
                                case 7: t7 = valor_1; break;
                                case 8: t8 = valor_1; break;
                                case 9: t9 = valor_1; break;
                                case 10: t10 = valor_1; break;
                                case 11: t11 = valor_1; break;
                                case 12: t12 = valor_1; break;
                                case 13: t13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(xTituloMarca, "Suma", Unidad, Ciudad_, xMercado, Unidad, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
        }

        public void Grupos_Tipo(string xCiudad, string xCodTipo, string xMercado, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2, string xMarcas_Grupo, string xTituloMarca)
        {
            if (xMoneda == 1)
            {
                Unidad = "PPU (ML)";
            }
            else
            {
                Unidad = "PPU (DOL.)";
            }

            if (xCiudad == "1")
            {
                Ciudad_ = "1. Capital";
            }
            else if (xCiudad == "1,2,5")
            {
                Ciudad_ = "0. Consolidado";
            }
            else { Ciudad_ = "2. Ciudades"; }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_PPU_TOTAL_TIPO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, xMarcas_Grupo);
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xCodTipo);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);

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

                            switch (i)
                            {
                                case 1: t1 = valor_1; break;
                                case 2: t2 = valor_1; break;
                                case 3: t3 = valor_1; break;
                                case 4: t4 = valor_1; break;
                                case 5: t5 = valor_1; break;
                                case 6: t6 = valor_1; break;
                                case 7: t7 = valor_1; break;
                                case 8: t8 = valor_1; break;
                                case 9: t9 = valor_1; break;
                                case 10: t10 = valor_1; break;
                                case 11: t11 = valor_1; break;
                                case 12: t12 = valor_1; break;
                                case 13: t13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(xTituloMarca, "Suma", Unidad, Ciudad_, xMercado, Unidad, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
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

        public void Hogares(string xCiudad)
        {
            Variable = "HOGARES";
            V2 = "";
            Recuperar_Marcas_Top_5_Retail_x_Hogar_Total(xCiudad, Codigo_cadena_Categoria);
            if (xCiudad == "1")
            {
                V1 = "Lima";
                Ciudad_ = "1. Capital";
                V1_ = "Cosmeticos";
                //Universo_Ciudad = 1;
            }
            else if (xCiudad == "1,2,5")
            {
                V1 = "Consolidado";
                Ciudad_ = "0. Consolidado";
                //Universo_Ciudad = 3;
            }
            else
            {
                V1 = "Ciudades";
                Ciudad_ = "2. Ciudades";
                V1_ = "Cosmeticos";
                //Universo_Ciudad = 2;
            }
            
            // TOTAL HOGARES NECESARIO PARA CALCULAR PENETRACION
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {                        
                        for (int i = 0; i < cols; i++)
                        {
                            if (xCiudad == "1")
                            {
                                Total_Mcdo_Hogar_Lima[0, i] = double.Parse(reader_1[i].ToString());
                                Total_Mcdo_Hogar_Temp[0, i] = double.Parse(reader_1[i].ToString());
                            }
                            else if (xCiudad == "2")
                            {
                                Total_Mcdo_Hogar_Ciudad[0, i] = double.Parse(reader_1[i].ToString());
                                Total_Mcdo_Hogar_Temp[0, i] = double.Parse(reader_1[i].ToString());
                            }
                            else
                            {
                                Total_Mcdo_Hogar_Pais[0, i] = double.Parse(reader_1[i].ToString());
                                Total_Mcdo_Hogar_Temp[0, i] = double.Parse(reader_1[i].ToString());
                            }
                        }                        
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()));
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));

                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD(V1_M, "Suma", Variable, Ciudad_, "0. Cosmeticos", "HOGARES", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                        //PENETRACIONES HOGAR TOTAL MERCADO
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Total_Mcdo_Hogar_Temp[0, 0] * 100);
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Total_Mcdo_Hogar_Temp[0, 1] * 100);
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Total_Mcdo_Hogar_Temp[0, 2] * 100);
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Total_Mcdo_Hogar_Temp[0, 3] * 100);
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Total_Mcdo_Hogar_Temp[0, 4] * 100);
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Total_Mcdo_Hogar_Temp[0, 5] * 100);
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Total_Mcdo_Hogar_Temp[0, 6] * 100);
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Total_Mcdo_Hogar_Temp[0, 7] * 100);
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Total_Mcdo_Hogar_Temp[0, 8] * 100);
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Total_Mcdo_Hogar_Temp[0, 9] * 100);
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Total_Mcdo_Hogar_Temp[0, 10] * 100);
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Total_Mcdo_Hogar_Temp[0, 11] * 100);
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Total_Mcdo_Hogar_Temp[0, 12] * 100);
                        Actualizar_BD_Penetracion(V1_M, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                    }
                }
            }     

            Grupo_Hogares(xCiudad, Codigo_Grupo_Belcorp_3M_Tipo, Codigo_cadena_Categoria, "00.Belcorp", Variable, Ciudad_, "0. Cosmeticos");
            Grupo_Hogares(xCiudad, Codigo_Grupo_Loreal_3M_Tipo, Codigo_cadena_Categoria, "50.Total L'Oreal", Variable, Ciudad_, "0. Cosmeticos");
            Grupo_Hogares(xCiudad, Codigo_Grupo_Lauder_3M_Tipo, Codigo_cadena_Categoria, "51.Total Estée Lauder", Variable, Ciudad_, "0. Cosmeticos");
 
            //HOGARES - CATEGORIA
            for (int x = 0; x < Codigo_Nombres_Categoria.GetLength(0); x++)
            {
                // HOGARES MERCADO POR CATEGORIA PARA CALCULO DE PENETRACION
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0].ToString());
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            for (int i = 1; i < cols; i++)
                            {
                                Total_Mcdo_Hogar_Temp[0, i - 1] = double.Parse(reader_1[i].ToString());
                            }
                        }                        
                    }
                }
                
                Recuperar_Marcas_Top_5_Retail_x_Hogar_Total(xCiudad, Codigo_Nombres_Categoria[x, 0].ToString()); ;
                Mercado = Codigo_Nombres_Categoria[x, 1];
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOTAL_MERCADO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {                                                  
                            t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()));
                            t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));

                            V1_M = Validar_marca(reader_1[0].ToString());
                            Actualizar_BD(V1_M, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                            //PENETRACION HOGARES POR CATEGORIA Y MARCAS
                            t1 = t1 / Total_Mcdo_Hogar_Temp[0, 0] * 100;
                            t2 = t2 / Total_Mcdo_Hogar_Temp[0, 1] * 100;
                            t3 = t3 / Total_Mcdo_Hogar_Temp[0, 2] * 100;
                            t4 = t4 / Total_Mcdo_Hogar_Temp[0, 3] * 100;
                            t5 = t5 / Total_Mcdo_Hogar_Temp[0, 4] * 100;
                            t6 = t6 / Total_Mcdo_Hogar_Temp[0, 5] * 100;
                            t7 = t7 / Total_Mcdo_Hogar_Temp[0, 6] * 100;
                            t8 = t8 / Total_Mcdo_Hogar_Temp[0, 7] * 100;
                            t9 = t9 / Total_Mcdo_Hogar_Temp[0, 8] * 100;
                            t10 = t10 / Total_Mcdo_Hogar_Temp[0, 9] * 100;
                            t11 = t11 / Total_Mcdo_Hogar_Temp[0, 10] * 100;
                            t12 = t12 / Total_Mcdo_Hogar_Temp[0, 11] * 100;
                            t13 = t13 / Total_Mcdo_Hogar_Temp[0, 12] * 100;
                            Actualizar_BD_Penetracion(V1_M, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                    Grupo_Hogares(xCiudad, Codigo_Grupo_Belcorp_3M_Tipo, Codigo_Nombres_Categoria[x, 0], "00.Belcorp", Variable, Ciudad_, Mercado);
                    Grupo_Hogares(xCiudad, Codigo_Grupo_Loreal_3M_Tipo, Codigo_Nombres_Categoria[x, 0], "50.Total L'Oreal", Variable, Ciudad_, Mercado);
                    Grupo_Hogares(xCiudad, Codigo_Grupo_Lauder_3M_Tipo, Codigo_Nombres_Categoria[x, 0], "51.Total Estée Lauder", Variable, Ciudad_, Mercado);
                }               
            }

            //HOGARES - TIPO
            for (int x = 0; x < Codigo_Nombres_Tipo.GetLength(0); x++)
            {
                // HOGARES MERCADO POR TIPO PARA CALCULO DE PENETRACION
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_HOGAR_TOTAL_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0].ToString());
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            for (int i = 1; i < cols; i++)
                            {
                                Total_Mcdo_Hogar_Temp[0, i - 1] = double.Parse(reader_1[i].ToString());
                            }
                        }
                    }
                }

                Recuperar_Marcas_Top_5_Retail_x_Hogar_Total_Tipo(xCiudad, Codigo_Nombres_Tipo[x, 0].ToString()); ;
                Mercado = Codigo_Nombres_Tipo[x, 1];
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOTAL_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {                 
                            t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()));
                            t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()));
                            t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()));
                            t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()));
                            t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()));
                            t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()));
                            t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()));
                            t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()));
                            t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()));
                            t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()));
                            t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()));
                            t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()));
                            t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()));

                            V1_M = Validar_marca(reader_1[0].ToString());
                            Actualizar_BD(V1_M, "Suma", Variable, Ciudad_, Mercado, "HOGARES", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                            //PENETRACION HOGARES POR CATEGORIA Y MARCAS
                            t1 = t1 / Total_Mcdo_Hogar_Temp[0, 0] * 100;
                            t2 = t2 / Total_Mcdo_Hogar_Temp[0, 1] * 100;
                            t3 = t3 / Total_Mcdo_Hogar_Temp[0, 2] * 100;
                            t4 = t4 / Total_Mcdo_Hogar_Temp[0, 3] * 100;
                            t5 = t5 / Total_Mcdo_Hogar_Temp[0, 4] * 100;
                            t6 = t6 / Total_Mcdo_Hogar_Temp[0, 5] * 100;
                            t7 = t7 / Total_Mcdo_Hogar_Temp[0, 6] * 100;
                            t8 = t8 / Total_Mcdo_Hogar_Temp[0, 7] * 100;
                            t9 = t9 / Total_Mcdo_Hogar_Temp[0, 8] * 100;
                            t10 = t10 / Total_Mcdo_Hogar_Temp[0, 9] * 100;
                            t11 = t11 / Total_Mcdo_Hogar_Temp[0, 10] * 100;
                            t12 = t12 / Total_Mcdo_Hogar_Temp[0, 11] * 100;
                            t13 = t13 / Total_Mcdo_Hogar_Temp[0, 12] * 100;
                            Actualizar_BD_Penetracion(V1_M, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                    Grupo_Hogares_Tipo(xCiudad, Codigo_Grupo_Belcorp_3M_Tipo, Codigo_Nombres_Tipo[x, 0], "00.Belcorp", Variable, Ciudad_, Mercado);
                    Grupo_Hogares_Tipo(xCiudad, Codigo_Grupo_Loreal_3M_Tipo, Codigo_Nombres_Tipo[x, 0], "50.Total L'Oreal", Variable, Ciudad_, Mercado);
                    Grupo_Hogares_Tipo(xCiudad, Codigo_Grupo_Lauder_3M_Tipo, Codigo_Nombres_Tipo[x, 0], "51.Total Estée Lauder", Variable, Ciudad_, Mercado);
                }
            }

        }

        public void Grupo_Hogares(string xCiudad, string xMarcaGrupo, string xCategoria, string xNombreGrupo, string xVariable, string xNombreCiudad, string xNombreMercado)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOTAL_MERCADO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, xMarcaGrupo);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xCategoria);

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

                            switch (i)
                            {
                                case 0: t1 = valor_1; break;
                                case 1: t2 = valor_1; break;
                                case 2: t3 = valor_1; break;
                                case 3: t4 = valor_1; break;
                                case 4: t5 = valor_1; break;
                                case 5: t6 = valor_1; break;
                                case 6: t7 = valor_1; break;
                                case 7: t8 = valor_1; break;
                                case 8: t9 = valor_1; break;
                                case 9: t10 = valor_1; break;
                                case 10: t11 = valor_1; break;
                                case 11: t12 = valor_1; break;
                                case 12: t13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(xNombreGrupo, "Suma", xVariable, xNombreCiudad, xNombreMercado, "HOGARES", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                        //if (xNombreMercado== "0. Cosmeticos")
                        //{
                            //PENETRACIONES HOGAR TOTAL MERCADO / CATEGORIA
                            t1 = t1 / Total_Mcdo_Hogar_Temp[0, 0] * 100;
                            t2 = t2 / Total_Mcdo_Hogar_Temp[0, 1] * 100;
                            t3 = t3 / Total_Mcdo_Hogar_Temp[0, 2] * 100;
                            t4 = t4 / Total_Mcdo_Hogar_Temp[0, 3] * 100;
                            t5 = t5 / Total_Mcdo_Hogar_Temp[0, 4] * 100;
                            t6 = t6 / Total_Mcdo_Hogar_Temp[0, 5] * 100;
                            t7 = t7 / Total_Mcdo_Hogar_Temp[0, 6] * 100;
                            t8 = t8 / Total_Mcdo_Hogar_Temp[0, 7] * 100;
                            t9 = t9 / Total_Mcdo_Hogar_Temp[0, 8] * 100;
                            t10 = t10 / Total_Mcdo_Hogar_Temp[0, 9] * 100;
                            t11 = t11 / Total_Mcdo_Hogar_Temp[0, 10] * 100;
                            t12 = t12 / Total_Mcdo_Hogar_Temp[0, 11] * 100;
                            t13 = t13 / Total_Mcdo_Hogar_Temp[0, 12] * 100;
                            Actualizar_BD_Penetracion(xNombreGrupo, "Suma", "PENETRACIONES", xNombreCiudad, xNombreMercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        //}
                        
                    }
                }
            }
        }
        public void Grupo_Hogares_Tipo(string xCiudad, string xMarcaGrupo, string xTipo, string xNombreGrupo, string xVariable, string xNombreCiudad, string xNombreMercado)
        {
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_HOGAR_TOTAL_TIPO_GRUPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, xMarcaGrupo);
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipo);

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

                            switch (i)
                            {
                                case 0: t1 = valor_1; break;
                                case 1: t2 = valor_1; break;
                                case 2: t3 = valor_1; break;
                                case 3: t4 = valor_1; break;
                                case 4: t5 = valor_1; break;
                                case 5: t6 = valor_1; break;
                                case 6: t7 = valor_1; break;
                                case 7: t8 = valor_1; break;
                                case 8: t9 = valor_1; break;
                                case 9: t10 = valor_1; break;
                                case 10: t11 = valor_1; break;
                                case 11: t12 = valor_1; break;
                                case 12: t13 = valor_1; break;
                            }
                        }
                        Actualizar_BD(xNombreGrupo, "Suma", xVariable, xNombreCiudad, xNombreMercado, "HOGARES", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                        //PENETRACIONES HOGAR TOTAL MERCADO / TIPO
                        t1 = t1 / Total_Mcdo_Hogar_Temp[0, 0] * 100;
                        t2 = t2 / Total_Mcdo_Hogar_Temp[0, 1] * 100;
                        t3 = t3 / Total_Mcdo_Hogar_Temp[0, 2] * 100;
                        t4 = t4 / Total_Mcdo_Hogar_Temp[0, 3] * 100;
                        t5 = t5 / Total_Mcdo_Hogar_Temp[0, 4] * 100;
                        t6 = t6 / Total_Mcdo_Hogar_Temp[0, 5] * 100;
                        t7 = t7 / Total_Mcdo_Hogar_Temp[0, 6] * 100;
                        t8 = t8 / Total_Mcdo_Hogar_Temp[0, 7] * 100;
                        t9 = t9 / Total_Mcdo_Hogar_Temp[0, 8] * 100;
                        t10 = t10 / Total_Mcdo_Hogar_Temp[0, 9] * 100;
                        t11 = t11 / Total_Mcdo_Hogar_Temp[0, 10] * 100;
                        t12 = t12 / Total_Mcdo_Hogar_Temp[0, 11] * 100;
                        t13 = t13 / Total_Mcdo_Hogar_Temp[0, 12] * 100;
                        Actualizar_BD_Penetracion(xNombreGrupo, "Suma", "PENETRACIONES", xNombreCiudad, xNombreMercado, "PENETRACIONES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
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
            if (x1.Length != (x2.Length / 2) / 5)
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
            HoraStart = DateTime.Now;
            Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_UU(xCiudad, 1);

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

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_MERCADO"))
            {
                //db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
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
                                periodo_Total_Mercado_UU[rows, i - 1] = valor_1;
                                Temp_Total_UU[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_UU[rows, i - 1] = valor_1;
                                Temp_Total_UU[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_UU[rows, i - 1] = valor_1;
                                Temp_Total_UU[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_UU"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
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
                //int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    //int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {              
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Total_UU[0, 0] * 100);
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Total_UU[0, 1] * 100);
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Total_UU[0, 2] * 100);
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Total_UU[0, 3] * 100);
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Total_UU[0, 4] * 100);
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Total_UU[0, 5] * 100);
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Total_UU[0, 6] * 100);
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Total_UU[0, 7] * 100);
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Total_UU[0, 8] * 100);
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Total_UU[0, 9] * 100);
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Total_UU[0, 10] * 100);
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Total_UU[0, 11] * 100);
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Total_UU[0, 12] * 100);
   
                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD_Penetracion(V1_M, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                        //rows++;
                    }
                }
            }           

            // BELCORP
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_UU[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_UU[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_UU[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_UU[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_UU[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_UU[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_UU[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_UU[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_UU[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_UU[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_UU[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_UU[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_UU[0, 12] * 100;

                        Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
            // LOREAL
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_UU[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_UU[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_UU[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_UU[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_UU[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_UU[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_UU[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_UU[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_UU[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_UU[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_UU[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_UU[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_UU[0, 12] * 100;

                        Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }
            // ESTEE LAUDER
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
            {
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
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

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = double.Parse(reader_1[1].ToString()) / Temp_Total_UU[0, 0] * 100;
                        t2 = double.Parse(reader_1[2].ToString()) / Temp_Total_UU[0, 1] * 100;
                        t3 = double.Parse(reader_1[3].ToString()) / Temp_Total_UU[0, 2] * 100;
                        t4 = double.Parse(reader_1[4].ToString()) / Temp_Total_UU[0, 3] * 100;
                        t5 = double.Parse(reader_1[5].ToString()) / Temp_Total_UU[0, 4] * 100;
                        t6 = double.Parse(reader_1[6].ToString()) / Temp_Total_UU[0, 5] * 100;
                        t7 = double.Parse(reader_1[7].ToString()) / Temp_Total_UU[0, 6] * 100;
                        t8 = double.Parse(reader_1[8].ToString()) / Temp_Total_UU[0, 7] * 100;
                        t9 = double.Parse(reader_1[9].ToString()) / Temp_Total_UU[0, 8] * 100;
                        t10 = double.Parse(reader_1[10].ToString()) / Temp_Total_UU[0, 9] * 100;
                        t11 = double.Parse(reader_1[11].ToString()) / Temp_Total_UU[0, 10] * 100;
                        t12 = double.Parse(reader_1[12].ToString()) / Temp_Total_UU[0, 11] * 100;
                        t13 = double.Parse(reader_1[13].ToString()) / Temp_Total_UU[0, 12] * 100;

                        Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                    }
                }
            }

            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Mercado UU : " + HoraFin.ToString());

            #region CATEGORIA
            HoraStart = DateTime.Now;
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
                                periodo_Total_Mercado_Categoria_UU[rows, i - 1] = valor_1;
                                Temp_Categoria_UU[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Categoria_UU[rows, i - 1] = valor_1;
                                Temp_Categoria_UU[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Categoria_UU[rows, i - 1] = valor_1;
                                Temp_Categoria_UU[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }

            int posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Categoria.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_Retail_x_Categoria_UU(xCiudad, int.Parse(Codigo_Nombres_Categoria[x, 0].ToString()), 1);

                for (int i = 0; i < Temp_Categoria_UU.Length; i++)
                {
                    if (Codigo_Nombres_Categoria[x, 0] == Temp_Categoria_UU[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Categoria[x, 1];
                        break;
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_CATEGORIA_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Categoria);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString())/ Temp_Categoria_UU[posicion, 1] * 100);
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString())/ Temp_Categoria_UU[posicion, 2] * 100);
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString())/ Temp_Categoria_UU[posicion, 3] * 100);
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString())/ Temp_Categoria_UU[posicion, 4] * 100);
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString())/ Temp_Categoria_UU[posicion, 5] * 100);
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString())/ Temp_Categoria_UU[posicion, 6] * 100);
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString())/ Temp_Categoria_UU[posicion, 7] * 100);
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString())/ Temp_Categoria_UU[posicion, 8] * 100);
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString())/ Temp_Categoria_UU[posicion, 9] * 100);
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString())/ Temp_Categoria_UU[posicion, 10] * 100);
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString())/ Temp_Categoria_UU[posicion, 11] * 100);
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString())/ Temp_Categoria_UU[posicion, 12] * 100);
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString())/ Temp_Categoria_UU[posicion, 13] * 100);                         
                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD_Penetracion(V1_M, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                            rows++;
                        }
                    }
                }

                // BELCORP
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = double.Parse(reader_1[1].ToString()) / Temp_Categoria_UU[posicion, 1] * 100;
                            t2 = double.Parse(reader_1[2].ToString()) / Temp_Categoria_UU[posicion, 2] * 100;
                            t3 = double.Parse(reader_1[3].ToString()) / Temp_Categoria_UU[posicion, 3] * 100;
                            t4 = double.Parse(reader_1[4].ToString()) / Temp_Categoria_UU[posicion, 4] * 100;
                            t5 = double.Parse(reader_1[5].ToString()) / Temp_Categoria_UU[posicion, 5] * 100;
                            t6 = double.Parse(reader_1[6].ToString()) / Temp_Categoria_UU[posicion, 6] * 100;
                            t7 = double.Parse(reader_1[7].ToString()) / Temp_Categoria_UU[posicion, 7] * 100;
                            t8 = double.Parse(reader_1[8].ToString()) / Temp_Categoria_UU[posicion, 8] * 100;
                            t9 = double.Parse(reader_1[9].ToString()) / Temp_Categoria_UU[posicion, 9] * 100;
                            t10 = double.Parse(reader_1[10].ToString()) / Temp_Categoria_UU[posicion, 10] * 100;
                            t11 = double.Parse(reader_1[11].ToString()) / Temp_Categoria_UU[posicion, 11] * 100;
                            t12 = double.Parse(reader_1[12].ToString()) / Temp_Categoria_UU[posicion, 12] * 100;
                            t13 = double.Parse(reader_1[13].ToString()) / Temp_Categoria_UU[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // LOREAL
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Categoria_UU[posicion, 1] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Categoria_UU[posicion, 2] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Categoria_UU[posicion, 3] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Categoria_UU[posicion, 4] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Categoria_UU[posicion, 5] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Categoria_UU[posicion, 6] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Categoria_UU[posicion, 7] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Categoria_UU[posicion, 8] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Categoria_UU[posicion, 9] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Categoria_UU[posicion, 10] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Categoria_UU[posicion, 11] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Categoria_UU[posicion, 12] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Categoria_UU[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // ESTEE LAUDER
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_MERCADO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_Nombres_Categoria[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Categoria_UU[0, 0] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Categoria_UU[0, 1] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Categoria_UU[0, 2] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Categoria_UU[0, 3] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Categoria_UU[0, 4] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Categoria_UU[0, 5] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Categoria_UU[0, 6] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Categoria_UU[0, 7] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Categoria_UU[0, 8] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Categoria_UU[0, 9] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Categoria_UU[0, 10] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Categoria_UU[0, 11] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Categoria_UU[0, 12] * 100;

                            Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }

            }

            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Categoria UU : " + HoraFin.ToString());
            #endregion CATEGORIA

            #region TIPO
            HoraStart = DateTime.Now;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODOS._SP_UNIDAD_TOTAL_TIPO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_cadena_Tipos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
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
                                periodo_Total_Mercado_Tipo_UU[rows, i - 1] = valor_1;
                                Temp_Tipo_UU[rows, i - 1] = valor_1;
                            }
                            else if (V1 == "Lima")
                            {
                                periodo_Total_Lima_Tipo_UU[rows, i - 1] = valor_1;
                                Temp_Tipo_UU[rows, i - 1] = valor_1;
                            }
                            else
                            {
                                periodo_Total_Ciudades_Tipo_UU[rows, i - 1] = valor_1;
                                Temp_Tipo_UU[rows, i - 1] = valor_1;
                            }
                        }
                        rows++;
                    }
                }
            }

            posicion = 0;
            for (int x = 0; x < Codigo_Nombres_Tipo.GetLength(0); x++)
            {
                Recuperar_Marcas_Top_5_Retail_x_Tipo(xCiudad, int.Parse(Codigo_Nombres_Tipo[x, 0].ToString()), 1);

                for (int i = 0; i < Temp_Tipo_UU.Length; i++)
                {
                    if (Codigo_Nombres_Tipo[x, 0] == Temp_Tipo_UU[i, 0].ToString())
                    {
                        posicion = i;
                        Mercado = Codigo_Nombres_Tipo[x, 1];
                        break;
                    }
                }
                //double[,] Temp_Tipo_Marca_Valor = new double[15, 15];  // TIPO MARCA VALORES

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);
                    int rows = 0;

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;

                        while (reader_1.Read())
                        {                    
                            t1 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString())/ Temp_Tipo_UU[posicion, 1] * 100);
                            t2 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString())/ Temp_Tipo_UU[posicion, 2] * 100);
                            t3 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString())/ Temp_Tipo_UU[posicion, 3] * 100);
                            t4 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString())/ Temp_Tipo_UU[posicion, 4] * 100);
                            t5 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString())/ Temp_Tipo_UU[posicion, 5] * 100);
                            t6 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString())/ Temp_Tipo_UU[posicion, 6] * 100);
                            t7 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString())/ Temp_Tipo_UU[posicion, 7] * 100);
                            t8 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString())/ Temp_Tipo_UU[posicion, 8] * 100);
                            t9 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString())/ Temp_Tipo_UU[posicion, 9] * 100);
                            t10 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString())/ Temp_Tipo_UU[posicion, 10] * 100);
                            t11 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString())/ Temp_Tipo_UU[posicion, 11] * 100);
                            t12 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString())/ Temp_Tipo_UU[posicion, 12] * 100);
                            t13 = (reader_1[14] == DBNull.Value ? 0 : double.Parse(reader_1[14].ToString())/ Temp_Tipo_UU[posicion, 13] * 100);

                            V1_M = Validar_marca(reader_1[1].ToString());
                            Actualizar_BD_Penetracion(V1_M, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);

                            rows++;
                        }
                    }
                }

                // BELCORP
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Belcorp_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
                    db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                    db_Zoho.AddInParameter(cmd_1, "_NSE", DbType.String, Codigo_cadena_NSE);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                    db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);
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
                    db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.Int32, xMoneda);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = double.Parse(reader_1[1].ToString()) / Temp_Tipo_UU[posicion, 1] * 100;
                            t2 = double.Parse(reader_1[2].ToString()) / Temp_Tipo_UU[posicion, 2] * 100;
                            t3 = double.Parse(reader_1[3].ToString()) / Temp_Tipo_UU[posicion, 3] * 100;
                            t4 = double.Parse(reader_1[4].ToString()) / Temp_Tipo_UU[posicion, 4] * 100;
                            t5 = double.Parse(reader_1[5].ToString()) / Temp_Tipo_UU[posicion, 5] * 100;
                            t6 = double.Parse(reader_1[6].ToString()) / Temp_Tipo_UU[posicion, 6] * 100;
                            t7 = double.Parse(reader_1[7].ToString()) / Temp_Tipo_UU[posicion, 7] * 100;
                            t8 = double.Parse(reader_1[8].ToString()) / Temp_Tipo_UU[posicion, 8] * 100;
                            t9 = double.Parse(reader_1[9].ToString()) / Temp_Tipo_UU[posicion, 9] * 100;
                            t10 = double.Parse(reader_1[10].ToString()) / Temp_Tipo_UU[posicion, 10] * 100;
                            t11 = double.Parse(reader_1[11].ToString()) / Temp_Tipo_UU[posicion, 11] * 100;
                            t12 = double.Parse(reader_1[12].ToString()) / Temp_Tipo_UU[posicion, 12] * 100;
                            t13 = double.Parse(reader_1[13].ToString()) / Temp_Tipo_UU[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("00.Belcorp", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // LOREAL
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Loreal_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Tipo_UU[posicion, 1] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Tipo_UU[posicion, 2] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Tipo_UU[posicion, 3] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Tipo_UU[posicion, 4] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Tipo_UU[posicion, 5] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Tipo_UU[posicion, 6] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Tipo_UU[posicion, 7] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Tipo_UU[posicion, 8] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Tipo_UU[posicion, 9] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Tipo_UU[posicion, 10] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Tipo_UU[posicion, 11] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Tipo_UU[posicion, 12] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Tipo_UU[posicion, 13] * 100;

                            Actualizar_BD_Penetracion("50.Total L'Oreal", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }
                // ESTEE LAUDER
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_TOTAL_TIPO_GRUPO_UU"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Grupo_Lauder_3M_Tipo);
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, Codigo_Nombres_Tipo[x, 0]);
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

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        while (reader_1.Read())
                        {
                            t1 = reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[1].ToString()) / Temp_Tipo_UU[0, 0] * 100;
                            t2 = reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[2].ToString()) / Temp_Tipo_UU[0, 1] * 100;
                            t3 = reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[3].ToString()) / Temp_Tipo_UU[0, 2] * 100;
                            t4 = reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[4].ToString()) / Temp_Tipo_UU[0, 3] * 100;
                            t5 = reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[5].ToString()) / Temp_Tipo_UU[0, 4] * 100;
                            t6 = reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[6].ToString()) / Temp_Tipo_UU[0, 5] * 100;
                            t7 = reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[7].ToString()) / Temp_Tipo_UU[0, 6] * 100;
                            t8 = reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[8].ToString()) / Temp_Tipo_UU[0, 7] * 100;
                            t9 = reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[9].ToString()) / Temp_Tipo_UU[0, 8] * 100;
                            t10 = reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[10].ToString()) / Temp_Tipo_UU[0, 9] * 100;
                            t11 = reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[11].ToString()) / Temp_Tipo_UU[0, 10] * 100;
                            t12 = reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[12].ToString()) / Temp_Tipo_UU[0, 11] * 100;
                            t13 = reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[13].ToString()) / Temp_Tipo_UU[0, 12] * 100;

                            Actualizar_BD_Penetracion("51.Total Estee Lauder", "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);
                        }
                    }
                }

            }
            HoraFin = DateTime.Now - HoraStart;
            Debug.Print("Total Tipo UU : " + HoraFin.ToString());
            #endregion TIPO



        }
        
        // GASTO MEDIO
        public void Gasto_Medio(string xCiudad, string xCab, string xAños, int xMoneda, string _PER12M_1, string _PER12M_2, string _PER6M_1, string _PER6M_2, string _PER3M_1, string _PER3M_2, string _PER1M_1, string _PER1M_2, string _PERYTDM_1, string _PERYTDM_2) 
        {
            Recuperar_Marcas_Top_5_Retail_x_GastoMedio_Total(xCiudad, _PER3M_2, Codigo_cadena_Categoria, xMoneda);

            if (xMoneda == 1)
            {
                Variable = "MONEDA LOCAL";
                VariableGM = "GASTO MEDIO (ML)";
                Variable_Promedio = "MONEDA LOCAL MES";                
            }
            else
            {
                Variable = "DOLARES";
                VariableGM = "GASTO MEDIO (DOL.)";
                Variable_Promedio = "DOLARES MES";               
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
         
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("PERIODO_MARCA._SP_GM_VALOR_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_CIUDAD", DbType.String, xCiudad);
                db_Zoho.AddInParameter(cmd_1, "_MARCAS", DbType.String, Codigo_Marcas_3M_Total_Cosmeticos);
                db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, Codigo_cadena_Categoria);                                
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, xCab);
                db_Zoho.AddInParameter(cmd_1, "_AÑOS", DbType.String, xAños);                
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
                db_Zoho.AddInParameter(cmd_1, "_MONEDA", DbType.String, xMoneda);
                //int rows = 0;

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    //int cols = reader_1.FieldCount;
                    while (reader_1.Read())
                    {
                        t1 = (reader_1[1] == DBNull.Value ? 0 : double.Parse(reader_1[15].ToString()) / double.Parse(reader_1[1].ToString()));
                        t2 = (reader_1[2] == DBNull.Value ? 0 : double.Parse(reader_1[16].ToString()) / double.Parse(reader_1[2].ToString()));
                        t3 = (reader_1[3] == DBNull.Value ? 0 : double.Parse(reader_1[17].ToString()) / double.Parse(reader_1[3].ToString()));
                        t4 = (reader_1[4] == DBNull.Value ? 0 : double.Parse(reader_1[18].ToString()) / double.Parse(reader_1[4].ToString()));
                        t5 = (reader_1[5] == DBNull.Value ? 0 : double.Parse(reader_1[19].ToString()) / double.Parse(reader_1[5].ToString()));
                        t6 = (reader_1[6] == DBNull.Value ? 0 : double.Parse(reader_1[20].ToString()) / double.Parse(reader_1[6].ToString()));
                        t7 = (reader_1[7] == DBNull.Value ? 0 : double.Parse(reader_1[21].ToString()) / double.Parse(reader_1[7].ToString()));
                        t8 = (reader_1[8] == DBNull.Value ? 0 : double.Parse(reader_1[22].ToString()) / double.Parse(reader_1[8].ToString()));
                        t9 = (reader_1[9] == DBNull.Value ? 0 : double.Parse(reader_1[23].ToString()) / double.Parse(reader_1[9].ToString()));
                        t10 = (reader_1[10] == DBNull.Value ? 0 : double.Parse(reader_1[24].ToString()) / double.Parse(reader_1[10].ToString()));
                        t11 = (reader_1[11] == DBNull.Value ? 0 : double.Parse(reader_1[25].ToString()) / double.Parse(reader_1[11].ToString()));
                        t12 = (reader_1[12] == DBNull.Value ? 0 : double.Parse(reader_1[26].ToString()) / double.Parse(reader_1[12].ToString()));
                        t13 = (reader_1[13] == DBNull.Value ? 0 : double.Parse(reader_1[27].ToString()) / double.Parse(reader_1[13].ToString()));

                        V1_M = Validar_marca(reader_1[0].ToString());
                        Actualizar_BD(V1_M, "Suma", VariableGM, Ciudad_, "0. Cosmeticos", VariableGM, "MENSUAL", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13);                        
                    }
                }
            }
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
                    Dato_1 = (_ANO_2 - _ANO_1);
                }
                if (_PER_12M_1 != 0)
                {
                    Dato_2 = (_PER_12M_2 - _PER_12M_1);
                }
                if (_PER_6M_1 != 0)
                {
                    Dato_3 = (_PER_6M_2 - _PER_6M_1);
                }
                if (_PER_3M_1 != 0)
                {
                    Dato_4 = (_PER_3M_2 - _PER_3M_1);
                }
                if (_PER_1M_1 != 0)
                {
                    Dato_5 = (_PER_1M_2 - _PER_1M_1);
                }
                if (_PER_YTD_1 != 0)
                {
                    Dato_6 = (_PER_YTD_2 - _PER_YTD_1);
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
