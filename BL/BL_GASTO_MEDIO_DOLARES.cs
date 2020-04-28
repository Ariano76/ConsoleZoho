using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;

namespace BL
{
    public class BL_Gasto_Medio_Dolares
    {
        public int[] sPeriodoActual = new int[2];
        private int[] Codigo_NSE = { 1, 2, 3, 4, 5 }; // NSE 
        private int[] Codigo_TIPOS = { 158, 161, 215, 202, 237, 226 }; // TIPOS 
        private string[] Cod_Categoria = { "123", "124", "127" }; // Categorias
        private string[] Codigo_Categoria = { "Fragancias", "Maquillaje", "Cuidado Personal" }; // Categorias
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor;
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_Valor_Region_NSE = new double[10, 48];  // VALORES POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Hogar_Region_NSE = new double[10, 48];  // HOGARES POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Valor_NSE_Region_Categoria = new double[10, 48];  // VALORES POR NSE, REGION Y CATEGORIA
        public double[,] sdata48Meses_x_Hogar_NSE_Region_Categoria = new double[10, 48];  // HOGARES POR NSE, REGION Y CATEGORIA
        public double[,] sdata48Meses_x_Valor_NSE_Region_Tipo = new double[10, 48];  // VALORES POR NSE, REGION Y TIPO
        public double[,] sdata48Meses_x_Hogar_NSE_Region_Tipo = new double[10, 48];  // HOGARES POR NSE, REGION Y TIPO
        public double[,] sdata48Meses_x_Valor_NSE_Categoria = new double[5, 48];  // VALORES POR NSE Y CATEGORIA
        public double[,] sdata48Meses_x_Hogar_NSE_Categoria = new double[5, 48];  // VALORES POR NSE Y CATEGORIA
        public double[,] sdata48Meses_x_Valor_NSE_Tipo = new double[5, 48];  // VALORES POR NSE Y TIPO
        public double[,] sdata48Meses_x_Hogar_NSE_Tipo = new double[5, 48];  // HOGARES POR NSE Y TIPO
        public double[,] sdata48Meses_x_Valor_NSE = new double[5, 48];  // VALORES POR NSE 
        public double[,] sdata48Meses_x_Hogar_NSE = new double[5, 48];  // VALORES POR NSE 
        public double[,] sdata48Meses_x_Valor_Tipo_Region_Modalidad = new double[4, 48];  // VALORES POR TIPO, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Tipo_Region_Modalidad = new double[4, 48];  // HOGARES POR TIPO, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Valor_Tipo_Region = new double[2, 48];  // VALORES POR TI PO Y REGION 
        public double[,] sdata48Meses_x_Hogar_Tipo_Region = new double[2, 48];  // HOGARES POR TIPO Y REGION 
        public double[,] sdata48Meses_x_Valor_Tipo_Modalidad = new double[2, 48];  // VALORES POR TIPO Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Tipo_Modalidad = new double[2, 48];  // HOGARES POR TIPO Y MODALIDAD
        public double[,] sdata48Meses_x_Valor_Tipo = new double[1, 48];  // VALORES POR TIPO
        public double[,] sdata48Meses_x_Hogar_Tipo = new double[1, 48];  // HOGARES POR TIPO
        public double[,] sdata48Meses_x_Valor_Region = new double[2, 48];  // VALORES POR REGION 
        public double[,] sdata48Meses_x_Hogar_Region = new double[2, 48];  // HOGARES POR REGION 
        public double[,] sdata48Meses_x_Valor = new double[1, 48];  // VALORES TOTAL
        public double[,] sdata48Meses_x_Hogar = new double[1, 48];  // HOGARES TOTAL
        public double[,] sdata48Meses_x_Valor_Categoria_Region_Modalidad = new double[4, 48];  // VALORES POR CATEGORIA, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Hogar_Categoria_Region_Modalidad = new double[4, 48];  // HOGARES POR CATEGORIA, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Valor_Categoria_Region = new double[2, 48];  // VALORES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Categoria_Region = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Valor_Categoria_Modalidad = new double[2, 48];  // VALORES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Categoria_Modalidad = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Valor_Categorias = new double[3, 48];  // VALORES POR CATEGORIA 
        public double[,] sdata48Meses_x_Hogar_Categorias = new double[3, 48];  // HOGARES POR CATEGORIA 
        public double[,] sdata48Meses_x_Valor_Modalidad_Region = new double[4, 48];  // VALORES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Modalidad_Region = new double[4, 48];  // HOGARES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Valor_Modalidad = new double[2, 48];  // VALORES POR CATEGORIA Y REGION 
        public double[,] sdata48Meses_x_Hogar_Modalidad = new double[2, 48];  // HOGARES POR CATEGORIA Y REGION 

        private readonly DateTime[] Periodos = new DateTime[7];
        string NSE_, NSE, V1, V1_, Mercado, Mercado_, Periodo, Ciudad_, xTipos_, Tipos_;
        byte fila = 0;
        public string resultadoBD;
        double valor_1, valor_2;
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

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_REGION_NSE"))
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
                            sdata48Meses_x_Valor_Region_NSE[rows, i - x] = valor_1;
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
                            sdata48Meses_x_Hogar_Region_NSE[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

            for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Valor_Region_NSE.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                    if (sdata48Meses_x_Valor_Region_NSE[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_Region_NSE[i, x] / sdata48Meses_x_Hogar_Region_NSE[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_NSE_REGION_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
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
                                sdata48Meses_x_Valor_NSE_Region_Categoria[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
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
                    for (int x = 0; x < sdata48Meses_x_Valor_NSE_Region_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                        if (sdata48Meses_x_Valor_NSE_Region_Categoria[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_NSE_Region_Categoria[i, x] / sdata48Meses_x_Hogar_NSE_Region_Categoria[i, x];
                        }

                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_NSE_REGION_TIPOS"))
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
                                sdata48Meses_x_Valor_NSE_Region_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                                sdata48Meses_x_Hogar_NSE_Region_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
                }
                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                if (contadorTotal == 10)
                {
                    for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Valor_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                            if (sdata48Meses_x_Valor_NSE_Region_Tipo[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Valor_NSE_Region_Tipo[i, x] / sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x];
                            }

                            Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Valor_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                            else if (i == 1)
                            {
                                NSE_ = "Medio";
                            }
                            else if (i == 2 | i == 6)
                            {
                                NSE_ = "Medio Bajo";
                            }
                            else if (i == 3 | i == 7)
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

                            if (sdata48Meses_x_Valor_NSE_Region_Tipo[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Valor_NSE_Region_Tipo[i, x] / sdata48Meses_x_Hogar_NSE_Region_Tipo[i, x];
                            }

                            Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
                sdata48Meses_x_Valor_NSE_Region_Tipo = new double[10, 48];
                sdata48Meses_x_Hogar_NSE_Region_Tipo = new double[10, 48];
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_NSE_CATEGORIA"))
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
                                sdata48Meses_x_Valor_NSE_Categoria[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

                for (int i = 0; i < 5; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_NSE_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                        if (sdata48Meses_x_Valor_NSE_Categoria[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_NSE_Categoria[i, x] / sdata48Meses_x_Hogar_NSE_Categoria[i, x];
                        }

                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_TIPO(string Cab, string xPeriodos)
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_NSE_TIPOS"))
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
                                sdata48Meses_x_Valor_NSE_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                                sdata48Meses_x_Hogar_NSE_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                    }
                }
                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_NSE_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_NSE_Tipo[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_NSE_Tipo[i, x] / sdata48Meses_x_Hogar_NSE_Tipo[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_NSE"))
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
                            sdata48Meses_x_Valor_NSE[rows, i - x] = valor_1;
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
                for (int x = 0; x < sdata48Meses_x_Valor_NSE.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                    if (sdata48Meses_x_Valor_NSE[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_NSE[i, x] / sdata48Meses_x_Hogar_NSE[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_TIPO_REGION_MODALIDAD"))
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
                                sdata48Meses_x_Valor_Tipo_Region_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Tipo_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Tipo_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Tipo_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Tipo_Region_Modalidad[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado_ = "01. Colonia Femeninas";
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado_ = "02. Colonia Masculinas";
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado_ = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado_ = "14. Roll-On";
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado_ = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_TOTAL_REGION_TIPOS"))
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
                                sdata48Meses_x_Valor_Tipo_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Tipo_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Tipo_Region[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Tipo_Region[i, x] / sdata48Meses_x_Hogar_Tipo_Region[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        /***/
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_TIPO_MODALIDAD"))
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
                                sdata48Meses_x_Valor_Tipo_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Tipo_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Tipo_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Tipo_Modalidad[i, x] / sdata48Meses_x_Hogar_Tipo_Modalidad[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_TIPOS)
            {
                switch (item)
                {
                    case 158:
                        NSE_ = "Colonia Femeninas";
                        Mercado_ = "01. Colonia Femeninas";
                        break;
                    case 161:
                        NSE_ = "Colonia Masculinas";
                        Mercado_ = "02. Colonia Masculinas";
                        break;
                    case 215:
                        NSE_ = "Humectante/Nutritiva Corporal";
                        Mercado_ = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        NSE_ = "Nutritiva Revit. Facial";
                        Mercado_ = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        NSE_ = "Roll-On";
                        Mercado_ = "14. Roll-On";
                        break;
                    case 226:
                        NSE_ = "Shampoo Adultos";
                        Mercado_ = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_TOTAL_TIPOS"))
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
                                sdata48Meses_x_Valor_Tipo[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        if (sdata48Meses_x_Valor_Tipo[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Tipo[i, x] / sdata48Meses_x_Hogar_Tipo[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //***///
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_REGION"))
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
                            sdata48Meses_x_Valor_Region[rows, i - x] = valor_1;
                            sdata48Meses_x_Valor[0, i - x] += valor_1;
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
                            sdata48Meses_x_Hogar[0, i - x] += valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

            for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Valor_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0)
                    {
                        NSE_ = "Lima";
                        Ciudad_ = "1. Capital";
                    }
                    else
                    {
                        NSE_ = "Ciudades";
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

                    if (sdata48Meses_x_Valor_Region[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_Region[i, x] / sdata48Meses_x_Hogar_Region[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD("Cosmeticos", "Suma", "GASTO MEDIO (DOL.)", Ciudad_, "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //***//
                    if (i == 0)
                    {
                        Actualizar_BD("Consolidado", "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, sdata48Meses_x_Valor[0, x] / sdata48Meses_x_Hogar[0, x], int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
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


                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_CATEG_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, item);
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
                                sdata48Meses_x_Valor_Categoria_Region_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Categoria_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Categoria_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Categoria_Region_Modalidad[i, x] / sdata48Meses_x_Hogar_Categoria_Region_Modalidad[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado, "GASTO MEDIO (DOL.)", "MENSUAL",
                            Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            foreach (var item in Codigo_Categoria)
            {
                switch (item)
                {
                    case "Fragancias":
                        NSE_ = "Fragancias";
                        Mercado_ = "1. Fragancias";
                        break;
                    case "Maquillaje":
                        NSE_ = "Maquillaje";
                        Mercado_ = "2. Maquillaje";
                        break;
                    case "Cuidado Personal":
                        NSE_ = "Cuidado Personal";
                        Mercado_ = "5. Cuidado Personal";
                        break;
                }


                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_CATEG_REGION"))
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
                                sdata48Meses_x_Valor_Categoria_Region[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Categoria_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Categoria_Region[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Categoria_Region[i, x] / sdata48Meses_x_Hogar_Categoria_Region[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(string Cab, string xPeriodos)
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


                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_CATEG_MODALIDAD"))
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
                                sdata48Meses_x_Valor_Categoria_Modalidad[rows, i - x] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }
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
                    }
                }

                /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
                for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Valor_Categoria_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                        if (sdata48Meses_x_Valor_Categoria_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Valor_Categoria_Modalidad[i, x] / sdata48Meses_x_Hogar_Categoria_Modalidad[i, x];
                        }
                        Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        //**//
                        Actualizar_BD(NSE, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIAS(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_CATEGORIAS"))
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
                            sdata48Meses_x_Valor_Categorias[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
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
                            sdata48Meses_x_Hogar_Categorias[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Valor_Categorias.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    switch (i)
                    {
                        case 0:
                            NSE_ = "Cuidado Personal";
                            Mercado_ = "5. Cuidado Personal";
                            break;
                        case 1:
                            NSE_ = "Fragancias";
                            Mercado_ = "1. Fragancias";
                            break;
                        case 2:
                            NSE_ = "Maquillaje";
                            Mercado_ = "2. Maquillaje";
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

                    if (sdata48Meses_x_Valor_Categorias[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_Categorias[i, x] / sdata48Meses_x_Hogar_Categorias[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_REGION_MODALIDAD"))
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
                            sdata48Meses_x_Valor_Modalidad_Region[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
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
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Valor_Modalidad_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                    if (sdata48Meses_x_Valor_Modalidad_Region[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_Modalidad_Region[i, x] / sdata48Meses_x_Hogar_Modalidad_Region[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL",
                        Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", Ciudad_, "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_VALOR_TOTAL_MODALIDAD"))
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
                            sdata48Meses_x_Valor_Modalidad[rows, i - x] = valor_1;
                        }
                        rows++;
                    }
                    contadorTotal = rows;
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
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */
            for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Valor_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0)
                    {
                        NSE_ = "VD";
                        Mercado_= "1. VD";
                    }
                    else
                    {
                        NSE_ = "VR";
                        Mercado_= "2. VR";
                    }

                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    if (sdata48Meses_x_Valor_Modalidad[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Valor_Modalidad[i, x] / sdata48Meses_x_Hogar_Modalidad[i, x];
                    }

                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", "0. Cosmeticos", "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    //**//
                    Actualizar_BD(NSE_, "Suma", "GASTO MEDIO (DOL.)", "0. Consolidado", Mercado_, "GASTO MEDIO (DOL.)", "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
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
