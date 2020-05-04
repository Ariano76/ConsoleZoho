using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;

namespace BL
{
    public class BL_Penetracion_Hogar
    {
        public int[] sPeriodoActual = new int[2];
        private int[] Codigo_NSE = { 1, 2, 3, 4, 5 }; // NSE 
        private int[] Codigo_TIPOS = { 158, 161, 215, 202, 237, 226 }; // TIPOS 
        private string[] Cod_Categoria = { "123", "124", "127" }; // Categorias
        private string[] Codigo_Categoria = { "Fragancias", "Maquillaje", "Cuidado Personal" }; // Categorias
        private string[] Categorias = { "Fragancias", "Maquillaje", "Tratamiento Facial", "Tratamiento Corporal", "Cuidado Personal" }; // Categorias
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor;
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_Hogar_Universo_Capital_Ciudad = new double[10, 48];  // VALORES POR NSE 
        public double[,] sdata48Meses_x_Hogar_Universo_Pais_NSE = new double[5, 48];  // VALORES POR NSE 
        public double[,] sdata48Meses_x_Hogar_Universo_Ciudad_Total = new double[2, 48];  // VALORES POR TOTAL CIUDAD
        public double[,] sdata48Meses_x_Hogar_Universo_Total = new double[1, 48];  // VALORES POR TOTAL CIUDAD

        public double[,] sdata48Meses_x_Hogar_NSE = new double[5, 48];  // VALORES POR NSE 
        public double[,] sdata48Meses_x_Hogar_NSE_Region = new double[10, 48];  // HOGARES POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Hogar_NSE_Region_Categoria = new double[10, 48];  // HOGARES POR NSE, REGION Y CATEGORIA        
        public double[,] sdata48Meses_x_Hogar_NSE_Categoria = new double[5, 48];  // VALORES POR NSE Y CATEGORIA
        public double[,] sdata48Meses_x_Hogar_NSE_Tipo = new double[5, 48];  // HOGARES POR NSE Y TIPO
        public double[,] sdata48Meses_x_Hogar_NSE_Region_Tipo = new double[10, 48];  // HOGARES POR NSE, REGION Y TIPO
        public double[,] sdata48Meses_x_Hogar_Region = new double[2, 48];  // HOGARES POR REGION 
        public double[,] sdata48Meses_x_Hogar_Modalidad_Region = new double[4, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Categoria_Region = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Categoria_Region_Modalidad = new double[4, 48];  // HOGARES POR CATEGORIA, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Tipo_Region = new double[2, 48];  // HOGARES POR TIPO Y REGION 
        public double[,] sdata48Meses_x_Hogar_Tipo_Region_Modalidad = new double[4, 48];  // HOGARES POR TIPO, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Categoria_Modalidad = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Categorias = new double[3, 48];  // HOGARES POR CATEGORIA 
        public double[,] sdata48Meses_x_Hogar_Tipo_Modalidad = new double[2, 48];  // HOGARES POR TIPO Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Tipo = new double[1, 48];  // HOGARES POR TIPO
        public double[,] sdata48Meses_x_Hogar = new double[1, 48];  // HOGARES TOTAL
        public double[,] sdata48Meses_x_Hogar_Modalidad = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Region_Modalidad = new double[4, 48];  // HOGARES POR REGION 
        public double[,] sdata48Meses_x_Hogar_Categorias_Regiones = new double[10, 48];  // HOGARES POR CATEGORIAS Y REGIONES FRAG, MAQ, CP
        public double[,] sdata48Meses_x_Hogar_Categorias_All = new double[5, 48];  // HOGARES POR CATEGORIA 



        private readonly DateTime[] Periodos = new DateTime[7];
        string NSE_, NSE, V1, V1_, Mercado, Mercado_, Periodo, Ciudad_, xTipos_, Tipos_;
        byte fila = 0;
        public string resultadoBD;
        double valor_1 = 0;
        double valor_2 = 0;
        int contadorTotal;

        //Database db = DatabaseFactory.CreateDatabase("SQL_BD_BIP");
        static DatabaseProviderFactory factory = new DatabaseProviderFactory();
        Database db = factory.Create("SQL_BD_BIP");
        Database db_Zoho = factory.Create("ZOHO");

        static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code, 
            // you can retrieve it from a configuration file.
            return "Data Source=PEDT0108243; Initial Catalog=BIP; User Id=Procesamiento; Password=P@ssw0rd";
        }


        public void Leer_Ultimos_48_Meses_NSE(string Cab, string xPeriodos, string xPerInicial, string xPerFinal)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_UNIVERSO_HOGAR_PAIS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_INICIO", DbType.String, xPerInicial);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_FIN", DbType.String, xPerFinal);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Universo_Pais_NSE[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_NSE[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

            for (int i = 0; i < 5; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Hogar_NSE.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0)
                    {
                        NSE_ = "Alto";
                    }
                    else if (i == 1)
                    {
                        NSE_ = "Medio";
                    }
                    else if (i == 2)
                    {
                        NSE_ = "Medio Bajo";
                    }
                    else if (i == 3)
                    {
                        NSE_ = "Bajo";
                    }
                    else
                    {
                        NSE_ = "Muy Bajo";
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_Hogar_NSE[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar_NSE[i, x] / sdata48Meses_x_Hogar_Universo_Pais_NSE[i, x] * 100;
                    }

                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CATEGORIA_CONSOLIDADO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_NSE[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        Mercado = "2. Maquillaje";
                        break;
                    case "Cuidado Personal":
                        Mercado = "5. Cuidado Personal";
                        break;
                }
               
                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_NSE_Categoria[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 | i == 5)
                        {
                            NSE_ = "Alto";
                        }
                        else if (i == 1 | i == 6)
                        {
                            NSE_ = "Medio";
                        }
                        else if (i == 2 | i == 7)
                        {
                            NSE_ = "Medio Bajo";
                        }
                        else if (i == 3 | i == 8)
                        {
                            NSE_ = "Bajo";
                        }
                        else
                        {
                            NSE_ = "Muy Bajo";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Hogar_NSE_Categoria[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_NSE_Categoria[i, x] / sdata48Meses_x_Hogar_NSE[i, x] * 100;
                        }

                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_TIPO_CONSOLIDADO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        Mercado_ = "01. Colonia Femeninas";
                        break;
                    case 161:
                        Mercado_ = "02. Colonia Masculinas";
                        break;
                    case 215:
                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        Mercado_ = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        Mercado_ = "14. Roll-On";
                        break;
                    case 226:
                        Mercado_ = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los valores.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_NSE_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 )
                            {
                                NSE_ = "Alto";
                            }
                        else if (i == 1 )
                            {
                                NSE_ = "Medio";
                            }
                        else if (i == 2 )
                            {
                                NSE_ = "Medio Bajo";
                            }
                        else if (i == 3 )
                            {
                                NSE_ = "Bajo";
                            }
                        else
                            {
                                NSE_ = "Muy Bajo";
                            }

                        if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                        else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                        if (sdata48Meses_x_Hogar_NSE_Tipo[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                        else
                            {
                                valor_1 = sdata48Meses_x_Hogar_NSE_Tipo[i, x] / sdata48Meses_x_Hogar_NSE[i, x] * 100;
                            }
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_NSE_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        Mercado = "2. Maquillaje";
                        break;
                    case "Cuidado Personal":
                        Mercado = "5. Cuidado Personal";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_REGION_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_NSE_Region_Categoria[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

                for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Region_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i < 5)
                        {
                            Ciudad_ = "1. Capital";
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                        }

                        if (i == 0 | i == 5)
                        {
                            NSE_ = "Alto";
                        }
                        else if (i == 1 | i == 6)
                        {
                            NSE_ = "Medio";
                        }
                        else if (i == 2 | i == 7)
                        {
                            NSE_ = "Medio Bajo";
                        }
                        else if (i == 3 | i == 8)
                        {
                            NSE_ = "Bajo";
                        }
                        else
                        {
                            NSE_ = "Muy Bajo";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Hogar_NSE_Region_Categoria[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_NSE_Region_Categoria[i, x] / sdata48Meses_x_Hogar_NSE_Region[i, x] * 100;
                        }

                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        Mercado_ = "01. Colonia Femeninas";
                        break;
                    case 161:
                        Mercado_ = "02. Colonia Masculinas";
                        break;
                    case 215:
                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        Mercado_ = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        Mercado_ = "14. Roll-On";
                        break;
                    case 226:
                        Mercado_ = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_REGION_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 2; // indicamos el numero de columna donde aparecen los valores.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_NSE_Region_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                if (contadorTotal == 10)
                {
                    for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (i < 5)
                            {
                                Ciudad_ = "1. Capital";
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                            }

                            if (i == 0 | i == 5)
                            {
                                NSE_ = "Alto";
                            }
                            else if (i == 1 | i == 6)
                            {
                                NSE_ = "Medio";
                            }
                            else if (i == 2 | i == 7)
                            {
                                NSE_ = "Medio Bajo";
                            }
                            else if (i == 3 | i == 8)
                            {
                                NSE_ = "Bajo";
                            }
                            else
                            {
                                NSE_ = "Muy Bajo";
                            }

                            if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }

                            if (sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x] / sdata48Meses_x_Hogar_NSE_Region[i, x] * 100;
                            }

                            Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (i < 5)
                            {
                                Ciudad_ = "1. Capital";
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                            }

                            switch (i)
                            {
                                case 0:
                                    fila = 0;
                                    NSE_ = "Alto";
                                    break;
                                case 1:
                                    fila = 1;
                                    NSE_ = "Medio";
                                    break;
                                case 2:
                                    fila = 2;
                                    NSE_ = "Medio Bajo";
                                    break;
                                case 3:
                                    fila = 3;
                                    NSE_ = "Bajo";
                                    break;
                                case 4:
                                    fila = 4;
                                    NSE_ = "Muy Bajo";
                                    break;
                                case 5:
                                    fila = 5;
                                    NSE_ = "Alto";
                                    break;
                                case 6:
                                    fila = 7;
                                    NSE_ = "Medio Bajo";
                                    break;
                                case 7:
                                    fila = 8;
                                    NSE_ = "Bajo";
                                    break;
                                case 8:
                                    fila = 9;
                                    NSE_ = "Muy Bajo";
                                    break;
                            }

                            if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }

                            if (sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x] / sdata48Meses_x_Hogar_NSE_Region[fila, x] * 100;
                            }

                            Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
                sdata48Meses_x_Hogar_NSE_Region_Tipo = new double[10, 48];
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD(string Cab, string xPeriodos, string xPerInicial, string xPerFinal)
        {
            /* ESTE PROCEDIMIENTO RECUPERA EL UNIVERSO DE HOGARES EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_UNIVERSO_HOGAR_CAPITAL_CIUDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_INICIO", DbType.String, xPerInicial);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_FIN", DbType.String, xPerFinal);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Universo_Capital_Ciudad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_NSE_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_NSE_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

            for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Hogar_NSE_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i < 5)
                    {
                        Ciudad_ = "1. Capital";
                    }
                    else
                    {
                        Ciudad_ = "2. Ciudades";
                    }

                    if (i == 0 | i == 5)
                    {
                        NSE_ = "Alto";
                    }
                    else if (i == 1 | i == 6)
                    {
                        NSE_ = "Medio";
                    }
                    else if (i == 2 | i == 7)
                    {
                        NSE_ = "Medio Bajo";
                    }
                    else if (i == 3 | i == 8)
                    {
                        NSE_ = "Bajo";
                    }
                    else
                    {
                        NSE_ = "Muy Bajo";
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_Hogar_NSE_Region[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar_NSE_Region[i, x] / sdata48Meses_x_Hogar_Universo_Capital_Ciudad[i, x] * 100;
                    }

                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }                    
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Modalidad_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Hogar_Modalidad_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0 | i == 2)
                    {
                        NSE_ = "VD";
                        Mercado_ = "1. VD";                        
                    }
                    else
                    {
                        NSE_ = "VR";
                        Mercado_ = "2. VR";
                    }

                    if (i == 0 | i == 1)
                    {
                        Ciudad_ = "1. Capital";
                        fila = 0;
                    }
                    else
                    {
                        Ciudad_ = "2. Ciudades";
                        fila = 1;
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    if (sdata48Meses_x_Hogar_Modalidad_Region[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar_Modalidad_Region[i, x] / sdata48Meses_x_Hogar_Region[fila, x] * 100;
                    }

                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                        Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double valor_1;

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado = "2. Maquillaje";
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado = "5. Cuidado Personal";
                        break;
                }


                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_REGION"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los valores.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }                        
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Region_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categoria_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 | i == 1)
                        {
                            Ciudad_ = "1. Capital";
                            fila = 0;
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                            fila = 1;
                        }

                        if (i == 0 | i == 2)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Categoria_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categoria_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Categoria_Region[fila, x] * 100;
                        }
                        //Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                        //    Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */            

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Region_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }


            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO_REGION"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los valores.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo_Region_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Tipo_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 | i == 1)
                        {
                            Ciudad_ = "1. Capital";
                            fila = 0;
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                            fila = 1;
                        }

                        if (i == 0 | i == 2)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Tipo_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                            valor_2 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Tipo_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Tipo_Region[fila, x] * 100;
                            valor_2 = sdata48Meses_x_Hogar_Tipo_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Region_Modalidad[i, x] * 100;
                        }
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_2, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIAS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Categorias[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado = "1. Fragancias";
                        fila = 1;
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado = "2. Maquillaje";
                        fila = 2;
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado = "5. Cuidado Personal";
                        fila = 0;
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categoria_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Categoria_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categoria_Modalidad[i, x] / sdata48Meses_x_Hogar_Categorias[fila, x] * 100;
                        }
                        //Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL",
                        //    Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "PENETRACIONES", "0. Consolidado", Mercado, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 0; // indicamos el numero de columna donde aparecen los valores.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }                        
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Tipo_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Tipo_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Tipo_Modalidad[i, x] / sdata48Meses_x_Hogar_Tipo[0, x] * 100;
                        }
                        //Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES", "MENSUAL",
                        //    Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "PENETRACIONES", "0. Consolidado", Mercado, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TOTAL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 0; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar[rows, i - x] = valor_1;
                        }
                        rows++;
                    }                    
                }
            }
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Hogar_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0)
                    {
                        NSE_ = "VD";
                        Mercado_ = "1. VD";
                    }
                    else
                    {
                        NSE_ = "VR";
                        Mercado_ = "2. VR";
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    if (sdata48Meses_x_Hogar_Modalidad[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar_Modalidad[i, x] / sdata48Meses_x_Hogar[0, x] * 100;
                    }

                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_REGION(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIAS_REGIONES"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Categorias_Regiones[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado_ = "01. Colonia Femeninas";
                        fila = 2;
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado_ = "02. Colonia Masculinas";
                        fila = 2;
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                        fila = 6;
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado_ = "08. Nutritiva Revit. Facial";
                        fila = 8;
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado_ = "14. Roll-On";
                        fila = 0;
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado_ = "10. Shampoo Adultos";
                        fila = 0;
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO_REGION"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Tipo_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            Ciudad_ = "1. Capital";
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Tipo_Region[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Tipo_Region[i, x] / sdata48Meses_x_Hogar_Categorias_Regiones[fila + i, x] * 100;
                        }
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //***///
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_MODALIDAD_CONSOLIDADO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Tipo_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Tipo_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Tipo_Modalidad[i, x] / sdata48Meses_x_Hogar_Modalidad[i, x] * 100;
                        }
                        //**//
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_CONSOLIDADO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIAS_ALL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Categorias_All[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        fila = 1;
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        fila = 1;
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        fila = 3;
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        fila = 4;
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado = "14. Roll-On";
                        fila = 0;
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        fila = 0;
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TIPO"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 0; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Tipo[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Tipo[i, x] / sdata48Meses_x_Hogar_Categorias_All[fila, x] * 100;
                        }
                        //**//
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 2; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Region_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado = "2. Maquillaje";
                        break;
                    case "Tratamiento Facial":
                        NSE_ = "Tratamiento Facial";
                        Mercado = "3. Tratamiento Facial";
                        break;
                    case "Tratamiento Corporal":
                        NSE_ = "Tratamiento Corporal";
                        Mercado = "4. Tratamiento Corporal";
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado = "5. Cuidado Personal";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 2; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Region_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categoria_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 | i == 1)
                        {
                            Ciudad_ = "1. Capital";
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                        }

                        if (i == 0 | i == 2)
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Categoria_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categoria_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Region_Modalidad[i, x] * 100;
                        }
                        //**//
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado_, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }                    
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado = "2. Maquillaje";
                        break;
                    case "Tratamiento Facial":
                        NSE_ = "Tratamiento Facial";
                        Mercado = "3. Tratamiento Facial";
                        break;
                    case "Tratamiento Corporal":
                        NSE_ = "Tratamiento Corporal";
                        Mercado = "4. Tratamiento Corporal";
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado = "5. Cuidado Personal";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_REGION"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categoria_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            Ciudad_ = "1. Capital";
                        }
                        else
                        {
                            Ciudad_ = "2. Ciudades";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Hogar_Categoria_Region[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categoria_Region[i, x] / sdata48Meses_x_Hogar_Region[i, x] * 100;
                        }

                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", Ciudad_, Mercado, "PENETRACIONES (%)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado = "2. Maquillaje";
                        break;
                    case "Tratamiento Facial":
                        NSE_ = "Tratamiento Facial";
                        Mercado = "3. Tratamiento Facial";
                        break;
                    case "Tratamiento Corporal":
                        NSE_ = "Tratamiento Corporal";
                        Mercado = "4. Tratamiento Corporal";
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado = "5. Cuidado Personal";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIA_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                        while (reader_1.Read())
                        {
                            for (int i = x; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Hogar_Categoria_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categoria_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0 )
                        {
                            Mercado_ = "1. VD";
                            NSE = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Categoria_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categoria_Modalidad[i, x] / sdata48Meses_x_Hogar_Modalidad[i, x] * 100;
                        }
                        //**//
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_CONSOLIDADO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */
            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TOTAL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 0; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }
            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_CATEGORIAS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Categorias_All[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar_Categorias_All.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (i == 0)
                        {
                            NSE_ = "Cuidado Personal";
                            Mercado_ = "5. Cuidado Personal";
                        }                        
                        else if (i == 1)
                        {
                            NSE_ = "Fragancias";
                            Mercado_ = "1. Fragancias";
                        }
                        else if (i == 2)
                        {
                            NSE_ = "Maquillaje";
                            Mercado_ = "2. Maquillaje";
                        }
                        else if (i == 3)
                        {
                            NSE_ = "Tratamiento Facial";
                            Mercado_ = "3. Tratamiento Facial";
                        }
                        else
                        {
                            NSE_ = "Tratamiento Corporal";
                            Mercado_ = "4. Tratamiento Corporal";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Hogar_Categorias_All[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Hogar_Categorias_All[i, x] / sdata48Meses_x_Hogar[0, x] * 100;
                        }
                        //**//
                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

                        Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", Mercado_, "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD(string Cab, string xPeriodos, string xPerInicial, string xPerFinal)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_UNIVERSO_HOGAR_CAPITAL_CIUDAD_TOTAL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_INICIO", DbType.String, xPerInicial);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_FIN", DbType.String, xPerFinal);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Universo_Ciudad_Total[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 1; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Hogar_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0 )
                    {
                        Ciudad_ = "1. Capital";
                        NSE_ = "Lima";
                    }
                    else
                    {
                        Ciudad_ = "2. Ciudades";
                        NSE_ = "Ciudades";
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    if (sdata48Meses_x_Hogar_Region[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar_Region[i, x] / sdata48Meses_x_Hogar_Universo_Ciudad_Total[i, x] * 100;
                    }

                    Actualizar_BD("Cosmeticos", "Suma", "PENETRACIONES", Ciudad_, "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //***
                    Actualizar_BD(NSE_, "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TOTAL_PAIS(string Cab, string xPeriodos, string xPerInicial, string xPerFinal)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_UNIVERSO_HOGAR_TOTAL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_INICIO", DbType.String, xPerInicial);
                db_Zoho.AddInParameter(cmd_1, "_PERIODO_FIN", DbType.String, xPerFinal);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 0; // indicamos el numero de columna donde aparecen los valores.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar_Universo_Total[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }
            
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_HOGAR_TOTAL"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    int x = 0; // indicamos el numero de columna donde aparecen los Hogares.
                    while (reader_1.Read())
                    {
                        for (int i = x; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Hogar[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
                }
            }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Hogar.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }                        
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    if (sdata48Meses_x_Hogar[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Hogar[i, x] / sdata48Meses_x_Hogar_Universo_Total[i, x] * 100;
                    }
                    Actualizar_BD("Consolidado", "Suma", "PENETRACIONES", "0. Consolidado", "0. Cosmeticos", "PENETRACIONES (%)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
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
