using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Data;
using System.Data.Common;

namespace BL
{
    public class BL_Share_Unidades
    {
        public int[] sPeriodoActual = new int[2];
        private int[] Codigo_NSE = { 1, 2, 3, 4, 5 }; // NSE 
        private int[] Codigo_TIPOS = { 158, 161, 215, 202, 237, 226 }; // TIPOS 
        private string[] Codigo_Categoria = { "Fragancias", "Maquillaje", "Cuidado Personal" }; // Categorias
        //public double[] sShareValor = new double[2000];

        public double[] sShareValor;
        public static readonly string[] sPeriodos48Meses = new string[2];
        public static readonly string[] sCabecera48Meses = new string[48];

        public double[,] sdata48Meses_x_Region_NSE = new double[10, 48];  // TOTAL POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Region_Total = new double[2, 48];  // TOTAL POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_Region_Modalidad = new double[4, 48];   // TOTAL POR REGION, MODALIDAD
        public double[,] sdata48Meses_x_Region_Categoria_Total = new double[2, 48];  // TOTAL POR REGION Y NSE VALORES
        public double[,] sdata48Meses_x_NSE_Region_Categoria = new double[10, 48];   // TOTAL POR NSE, REGION Y CATEGORIA
        public double[,] sdata48Meses_x_NSE_Region_Tipo = new double[10, 48];   // TOTAL POR NSE, REGION Y TIPO
        public double[,] sdata48Meses_x_Region_Tipo_Total = new double[2, 48];  // TOTAL POR REGION Y TIPO VALORES
        public double[,] sdata48Meses_x_Categoria_Total = new double[1, 48];  // TOTAL POR CATEGORIA
        public double[,] sdata48Meses_x_Categoria_Region_Total = new double[1, 48];  // TOTAL POR CATEGORIA y REGION
        public double[,] sdata48Meses_x_NSE_Categoria = new double[5, 48];   // TOTAL POR NSE Y CATEGORIA
        public double[,] sdata48Meses_x_Tipo_Total = new double[1, 48];  // TOTAL POR TIPO VALORES
        public double[,] sdata48Meses_x_NSE_Tipo = new double[5, 48];   // TOTAL POR NSE Y TIPO
        public double[,] sdata48Meses_x_NSE = new double[5, 48];  // TOTAL POR NSE
        public double[,] sdata48Meses_x_TOTAL = new double[1, 48];  // TOTAL MERCADO
        public double[,] sdata48Meses_x_Ciudad = new double[2, 48];  // TOTAL POR NSE
        public double[,] sdata48Meses_x_Categoria_Region_Modalidad = new double[4, 48];   // TOTAL POR CATEGORIA, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Categoria_Region = new double[2, 48];   // TOTAL POR CATEGORIA Y MODALIDAD
        public double[,] sdata48Meses_x_Categoria_Modalidad = new double[2, 48];   // TOTAL POR CATEGORIA Y MODALIDAD        
        public double[,] sdata48Meses_x_Modalidad_Total = new double[2, 48];          // TOTAL POR MODALIDAD
        public double[,] sdata48Meses_x_Categoria = new double[3, 48];   // TOTAL POR CATEGORIA                
        public double[,] sdata48Meses_x_Tipo_Region_Modalidad = new double[4, 48];   // TOTAL POR TIPO, REGION Y MODALIDAD
        public double[,] sdata48Meses_x_Tipo_Region = new double[2, 48];   // TOTAL POR TIPO Y REGION 
        public double[,] sdata48Meses_x_Tipo_Modalidad = new double[2, 48];   // TOTAL POR TIPO Y MODALIDAD

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

        public void Leer_Ultimos_48_Meses_CIUDAD_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR REGIONES Y NSE (CAPITAL Y PROVINCIA Y NSE) */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
                        for (int i = 1; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Region_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_REGION_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
                        for (int i = 2; i < cols; i++)
                        {
                            if (reader_1[i] == DBNull.Value)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = double.Parse(reader_1[i].ToString());
                            }
                            sdata48Meses_x_Region_NSE[rows, i - 2] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* INSERTANDO VALORES - PPU REGION Y NSE A BD */

            for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Region_NSE.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i < 5)
                    {
                        Mercado_ = "1. Capital";
                        fila = 0;
                        switch (i)
                        {
                            case 0:
                                NSE_ = "Alto";
                                break;
                            case 1:
                                NSE_ = "Medio";
                                break;
                            case 2:
                                NSE_ = "Medio Bajo";
                                break;
                            case 3:
                                NSE_ = "Bajo";
                                break;
                            case 4:
                                NSE_ = "Muy Bajo";
                                break;
                        }
                    }
                    else
                    {
                        Mercado_ = "2. Ciudades";
                        fila = 1;
                        switch (i)
                        {
                            case 5:
                                NSE_ = "Alto";
                                break;
                            case 6:
                                NSE_ = "Medio";
                                break;
                            case 7:
                                NSE_ = "Medio Bajo";
                                break;
                            case 8:
                                NSE_ = "Bajo";
                                break;
                            case 9:
                                NSE_ = "Muy Bajo";
                                break;
                        }
                    }
                    if (x < 9)
                    {
                        Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }
                    else
                    {
                        Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                    }

                    if (sdata48Meses_x_Region_NSE[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Region_NSE[i, x] / sdata48Meses_x_Region_Total[fila, x] * 100;
                    }

                    Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Mercado_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL",
                        Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR CIUDAD Y MODADLIDAD */

            double valor_1;

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
                    while (reader_1.Read())
                    {
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
                            sdata48Meses_x_Region_Modalidad[rows, i - 2] = valor_1;
                        }
                        rows++;
                    }
                }

                /* INSERTANDO VALORES - SHARE REGION Y MODALIDAD A BD */
                for (int i = 0; i < 4; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    for (int x = 0; x < sdata48Meses_x_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                        if (i % 2 == 0)
                        {
                            Mercado_ = "1. VD";
                            NSE_ = "VD";
                        }
                        else
                        {
                            Mercado_ = "2. VR";
                            NSE_ = "VR";
                        }

                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Region_Modalidad[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Region_Modalidad[i, x] / sdata48Meses_x_Region_Total[fila, x] * 100;
                        }

                        Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)",
                            "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            string xNSE = "";

            foreach (var item in Codigo_Categoria)
            {
                xNSE = item.ToString();
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
                            for (int i = 1; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Region_Categoria_Total[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_NSE_REGION_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
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
                                sdata48Meses_x_NSE_Region_Categoria[rows, i - 2] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 10; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_NSE_Region_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (i < 5)
                            {
                                Ciudad_ = "1. Capital";
                                fila = 0;
                            }
                            else
                            {
                                Ciudad_ = "2. Ciudades";
                                fila = 1;
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

                            if (sdata48Meses_x_NSE_Region_Categoria[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_NSE_Region_Categoria[i, x] / sdata48Meses_x_Region_Categoria_Total[fila, x] * 100;
                            }

                            Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL",
                                Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y TIPOS */
            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
                            for (int i = 1; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Region_Tipo_Total[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_NSE_REGION_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

                        while (reader_1.Read())
                        {
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
                                sdata48Meses_x_NSE_Region_Tipo[rows, i - 2] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    if (contadorTotal == 10)
                    {
                        for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                        {
                            for (int x = 0; x < sdata48Meses_x_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                            {
                                if (i < 5)
                                {
                                    Ciudad_ = "1. Capital";
                                    fila = 0;
                                }
                                else
                                {
                                    Ciudad_ = "2. Ciudades";
                                    fila = 1;
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

                                if (sdata48Meses_x_NSE_Region_Tipo[i, x] <= 0)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = sdata48Meses_x_NSE_Region_Tipo[i, x] / sdata48Meses_x_Region_Tipo_Total[fila, x] * 100;
                                }

                                Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                        {
                            for (int x = 0; x < sdata48Meses_x_NSE_Region_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                            {
                                if (i < 5)
                                {
                                    Ciudad_ = "1. Capital";
                                    fila = 0;
                                }
                                else
                                {
                                    Ciudad_ = "2. Ciudades";
                                    fila = 1;
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

                                if (sdata48Meses_x_NSE_Region_Tipo[i, x] <= 0)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = sdata48Meses_x_NSE_Region_Tipo[i, x] / sdata48Meses_x_Region_Tipo_Total[fila, x] * 100;
                                }

                                Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            }
                        }
                    }
                    sdata48Meses_x_NSE_Region_Tipo = new double[10, 48];
                    sdata48Meses_x_Region_Tipo_Total = new double[2, 48];
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            string xNSE = "";

            foreach (var item in Codigo_Categoria)
            {
                xNSE = item.ToString();
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
                            for (int i = 0; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Categoria_Total[rows, i] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_NSE_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
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
                                sdata48Meses_x_NSE_Categoria[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 5; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_NSE_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                            if (sdata48Meses_x_NSE_Categoria[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_NSE_Categoria[i, x] / sdata48Meses_x_Categoria_Total[0, x] * 100;
                            }

                            Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", Mercado, "UNIDADES (%)", "MENSUAL",
                                Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE_TIPO(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y TIPOS */
            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
                            for (int i = 0; i < cols; i++)
                            {
                                if (reader_1[i] == DBNull.Value)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = double.Parse(reader_1[i].ToString());
                                }
                                sdata48Meses_x_Tipo_Total[rows, i] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_NSE_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_NSE_Tipo[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                        contadorTotal = rows;
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    if (contadorTotal == 10)
                    {
                        for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                        {
                            for (int x = 0; x < sdata48Meses_x_NSE_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                                if (sdata48Meses_x_NSE_Tipo[i, x] <= 0)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = sdata48Meses_x_NSE_Tipo[i, x] / sdata48Meses_x_Tipo_Total[0, x] * 100;
                                }

                                Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < contadorTotal; i++) // LEYENDO LAS FILAS DEL ARRAY
                        {
                            for (int x = 0; x < sdata48Meses_x_NSE_Tipo.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                            {
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

                                if (sdata48Meses_x_NSE_Tipo[i, x] <= 0)
                                {
                                    valor_1 = 0;
                                }
                                else
                                {
                                    valor_1 = sdata48Meses_x_NSE_Tipo[i, x] / sdata48Meses_x_Tipo_Total[0, x] * 100;
                                }

                                Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            }
                        }
                    }
                    sdata48Meses_x_NSE_Tipo = new double[10, 48];
                    sdata48Meses_x_Tipo_Total = new double[2, 48];
                }
            }
        }

        public void Leer_Ultimos_48_Meses_NSE(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_NSE"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_NSE[rows, i - 1] = valor_1;
                            sdata48Meses_x_TOTAL[0, i - 1] += valor_1;
                        }
                        rows++;
                    }
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
            for (int i = 0; i < 5; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_NSE.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                    if (sdata48Meses_x_NSE[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_NSE[i, x] / sdata48Meses_x_TOTAL[0, x] * 100;
                    }
                    Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL",
                        Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CIUDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);
                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_Ciudad[rows, i - 1] = valor_1;
                            //sdata48Meses_x_TOTAL[0, i - 1] += valor_1;
                        }
                        rows++;
                    }
                }
            }
            /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
            for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
            {
                for (int x = 0; x < sdata48Meses_x_Ciudad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                {
                    if (i == 0)
                    {
                        NSE_ = "Lima";
                    }
                    else
                    {
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
                    if (sdata48Meses_x_Ciudad[i, x] <= 0)
                    {
                        valor_1 = 0;
                    }
                    else
                    {
                        valor_1 = sdata48Meses_x_Ciudad[i, x] / sdata48Meses_x_TOTAL[0, x] * 100;
                    }
                    Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL",
                        Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            string xNSE = "";

            foreach (var item in Codigo_Categoria)
            {
                xNSE = item.ToString();
                NSE_ = xNSE;
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_CATEG_REGION"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
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
                                sdata48Meses_x_Categoria_Region[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_CATEG_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
                        while (reader_1.Read())
                        {
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
                                sdata48Meses_x_Categoria_Region_Modalidad[rows, i - 2] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU NSE Y CATEGORIA A BD */
                    for (int i = 0; i < 4; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Categoria_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                                V1 = "VD";
                                NSE = "VD";
                            }
                            else
                            {
                                Mercado_ = "2. VR";
                                V1 = "VR";
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

                            if (sdata48Meses_x_Categoria_Region_Modalidad[i, x] <= 0)
                            {
                                valor_1 = 0;
                                valor_2 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Categoria_Region_Modalidad[i, x] / sdata48Meses_x_Region_Modalidad[i, x] * 100;
                                valor_2 = sdata48Meses_x_Categoria_Region_Modalidad[i, x] / sdata48Meses_x_Categoria_Region[fila, x] * 100;
                            }

                            Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)",
                                "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));

                            // RESULTADOS MODALIDAD - REGION - CATEGORIA
                            Actualizar_BD(NSE, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)",
                                "MENSUAL", Periodo, valor_2, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_REGION(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            string xNSE = "";

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_Region_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                xNSE = item.ToString();
                NSE_ = xNSE;
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
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
                                sdata48Meses_x_Categoria_Region[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU NSE Y CATEGORIA A BD */
                    byte fila;
                    for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Categoria_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (i == 0)
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

                            if (sdata48Meses_x_Categoria_Region[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Categoria_Region[i, x] / sdata48Meses_x_Region_Total[fila, x] * 100;
                            }
                            Actualizar_BD(NSE_, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)",
                                "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            string xNSE = "";

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_Modalidad_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_Categoria)
            {
                xNSE = item.ToString();
                NSE_ = xNSE;
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

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_CATEGORIA"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
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
                                sdata48Meses_x_Categoria_Total[rows, i] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_CATEG_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_CATEGORIA", DbType.String, xNSE);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;
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
                                sdata48Meses_x_Categoria_Modalidad[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU MODALIDAD Y CATEGORIA A BD */
                    for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Categoria_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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

                            if (sdata48Meses_x_Categoria_Modalidad[i, x] <= 0)
                            {
                                valor_1 = 0;
                                valor_2 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Categoria_Modalidad[i, x] / sdata48Meses_x_Modalidad_Total[i, x] * 100;
                                valor_2 = sdata48Meses_x_Categoria_Modalidad[i, x] / sdata48Meses_x_Categoria_Total[0, x] * 100;
                            }
                            Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", Mercado_, "UNIDADES (%)",
                                "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            // RESULTADO POR CATEGORIA, MODALIDAD CONSOLIDADO
                            Actualizar_BD(NSE, "%", "SHARE UNIDADES", "0. Consolidado", Mercado, "UNIDADES (%)",
                                "MENSUAL", Periodo, valor_2, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_CATEGORIA(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_TOTAL[rows, i] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_CATEGORIAS"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_Categoria[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }

                /* INSERTANDO VALORES - PPU MODALIDAD Y CATEGORIA A BD */
                for (int i = 0; i < 3; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    switch (i)
                    {
                        case 0:
                            NSE_ = "Cuidado Personal";
                            break;
                        case 1:
                            NSE_ = "Fragancias";
                            break;
                        case 2:
                            NSE_ = "Maquillaje";
                            break;
                    }
                    for (int x = 0; x < sdata48Meses_x_Categoria.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Categoria[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Categoria[i, x] / sdata48Meses_x_TOTAL[0, x] * 100;
                        }
                        Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)",
                            "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(string Cab, string xPeriodos)
        {

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_REGION_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;

                    while (reader_1.Read())
                    {
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
                            sdata48Meses_x_Region_Modalidad[rows, i - 2] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR TIPO, REGION Y MODALIDAD */
            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
                switch (item)
                {
                    case 158:
                        V1 = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        V1 = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        V1 = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        V1 = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        V1 = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        V1 = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_Tipo_Region[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TIPO_REGION_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

                        while (reader_1.Read())
                        {
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
                                sdata48Meses_x_Tipo_Region_Modalidad[rows, i - 2] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 4; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Tipo_Region_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                                V1_ = "VD";
                            }
                            else
                            {
                                Mercado_ = "2. VR";
                                V1_ = "VR";
                            }

                            if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }

                            if (sdata48Meses_x_Tipo_Region_Modalidad[i, x] <= 0)
                            {
                                valor_1 = 0;
                                valor_2 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Tipo_Region_Modalidad[i, x] / sdata48Meses_x_Region_Modalidad[i, x] * 100;
                                valor_2 = sdata48Meses_x_Tipo_Region_Modalidad[i, x] / sdata48Meses_x_Tipo_Region[fila, x] * 100;
                            }
                            Actualizar_BD(V1, "%", "SHARE UNIDADES", Ciudad_, Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            //******
                            Actualizar_BD(V1_, "%", "SHARE UNIDADES", Ciudad_, Mercado, "UNIDADES (%)", "MENSUAL", Periodo, valor_2,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_REGION(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR TIPO, REGION Y MODALIDAD */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;

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
                            sdata48Meses_x_Region_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
                switch (item)
                {
                    case 158:
                        V1 = "Colonia Femeninas";
                        break;
                    case 161:
                        V1 = "Colonia Masculinas";
                        break;
                    case 215:
                        V1 = "Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        V1 = "Nutritiva Revit. Facial";
                        break;
                    case 237:
                        V1 = "Roll-On";
                        break;
                    case 226:
                        V1 = "Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_REGION_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_Tipo_Region[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Tipo_Region.GetLength(1); x++) //LEYENDO LAS COLUMNAS
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
                            if (sdata48Meses_x_Tipo_Region[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Tipo_Region[i, x] / sdata48Meses_x_Region_Total[i, x] * 100;
                            }
                            Actualizar_BD(V1, "%", "SHARE UNIDADES", Ciudad_, "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR TIPO, REGION Y MODALIDAD */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;

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
                            sdata48Meses_x_Modalidad_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
                switch (item)
                {
                    case 158:
                        V1 = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        V1 = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        V1 = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        V1 = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        V1 = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        V1 = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_Tipo_Total[rows, i] = valor_1;
                            }
                            rows++;
                        }
                    }
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TIPO_MODALIDAD"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_Tipo_Modalidad[rows, i - 1] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Tipo_Modalidad.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (i == 0)
                            {
                                Mercado_ = "1. VD";
                                V1_ = "VD";
                            }
                            else
                            {
                                Mercado_ = "2. VR";
                                V1_ = "VR";
                            }

                            if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            if (sdata48Meses_x_Tipo_Modalidad[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Tipo_Modalidad[i, x] / sdata48Meses_x_Modalidad_Total[i, x] * 100;
                                valor_2 = sdata48Meses_x_Tipo_Modalidad[i, x] / sdata48Meses_x_Tipo_Total[0, x] * 100;
                            }
                            Actualizar_BD(V1, "%", "SHARE UNIDADES", "0. Consolidado", Mercado_, "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            //******
                            Actualizar_BD(V1_, "%", "SHARE UNIDADES", "0. Consolidado", Mercado, "UNIDADES (%)", "MENSUAL", Periodo, valor_2,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_TIPO_TOTAL(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR TIPO, REGION Y MODALIDAD */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;

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
                            sdata48Meses_x_TOTAL[rows, i] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            foreach (var item in Codigo_TIPOS)
            {
                xTipos_ = item.ToString();
                switch (item)
                {
                    case 158:
                        V1 = "Colonia Femeninas";
                        Mercado = "01. Colonia Femeninas";
                        break;
                    case 161:
                        V1 = "Colonia Masculinas";
                        Mercado = "02. Colonia Masculinas";
                        break;
                    case 215:
                        V1 = "Humectante/Nutritiva Corporal";
                        Mercado = "09. Humectante/Nutritiva Corporal";
                        break;
                    case 202:
                        V1 = "Nutritiva Revit. Facial";
                        Mercado = "08. Nutritiva Revit. Facial";
                        break;
                    case 237:
                        V1 = "Roll-On";
                        Mercado = "14. Roll-On";
                        break;
                    case 226:
                        V1 = "Shampoo Adultos";
                        Mercado = "10. Shampoo Adultos";
                        break;
                }

                using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_TIPOS"))
                {
                    db_Zoho.AddInParameter(cmd_1, "_TIPO", DbType.String, xTipos_);
                    db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                    db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                    using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                    {
                        int cols = reader_1.FieldCount;
                        int rows = 0;

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
                                sdata48Meses_x_Tipo_Total[rows, i] = valor_1;
                            }
                            rows++;
                        }
                    }

                    /* INSERTANDO VALORES - PPU REGION Y MODALIDAD A BD */
                    for (int i = 0; i < 1; i++) // LEYENDO LAS FILAS DEL ARRAY
                    {
                        for (int x = 0; x < sdata48Meses_x_Tipo_Total.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                        {
                            if (x < 9)
                            {
                                Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            else
                            {
                                Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                            }
                            if (sdata48Meses_x_Tipo_Total[i, x] <= 0)
                            {
                                valor_1 = 0;
                            }
                            else
                            {
                                valor_1 = sdata48Meses_x_Tipo_Total[i, x] / sdata48Meses_x_TOTAL[0, x] * 100;
                            }
                            Actualizar_BD(V1, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)", "MENSUAL", Periodo, valor_1,
                                int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                            //******
                            //Actualizar_BD(V1_, "SUMA", "PPU (DOL)", Ciudad_, Mercado, "PPU (DOL)", "MENSUAL", Periodo, valor_1,
                            //    int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
                        }
                    }
                }
            }
        }

        public void Leer_Ultimos_48_Meses_MODALIDAD(string Cab, string xPeriodos)
        {
            /* ESTE PROCEDIMIENTO RECUPERA LA DATA EN FORMATO PIVOT POR NSE, CIUDAD Y CATEGORIA */
            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MERCADO"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_TOTAL[rows, i] = valor_1;
                        }
                        rows++;
                    }
                }
            }

            using (DbCommand cmd_1 = db_Zoho.GetStoredProcCommand("_SP_SHARE_UNIDAD_TOTAL_MODALIDAD"))
            {
                db_Zoho.AddInParameter(cmd_1, "_PERIODO", DbType.String, xPeriodos);
                db_Zoho.AddInParameter(cmd_1, "_CABECERA", DbType.String, Cab);

                using (IDataReader reader_1 = db_Zoho.ExecuteReader(cmd_1))
                {
                    int cols = reader_1.FieldCount;
                    int rows = 0;
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
                            sdata48Meses_x_Modalidad_Total[rows, i - 1] = valor_1;
                        }
                        rows++;
                    }
                }

                /* INSERTANDO VALORES - PPU MODALIDAD Y CATEGORIA A BD */
                for (int i = 0; i < 2; i++) // LEYENDO LAS FILAS DEL ARRAY
                {
                    switch (i)
                    {
                        case 0:
                            NSE_ = "VD";
                            break;
                        case 1:
                            NSE_ = "VR";
                            break;
                    }
                    for (int x = 0; x < sdata48Meses_x_Modalidad_Total.GetLength(1); x++) //LEYENDO LAS COLUMNAS
                    {
                        if (x < 9)
                        {
                            Periodo = "0" + (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }
                        else
                        {
                            Periodo = (x + 1) + ". " + BD_Zoho.sCabecera48Meses[x];
                        }

                        if (sdata48Meses_x_Modalidad_Total[i, x] <= 0)
                        {
                            valor_1 = 0;
                        }
                        else
                        {
                            valor_1 = sdata48Meses_x_Modalidad_Total[i, x] / sdata48Meses_x_TOTAL[0, x] * 100;
                        }
                        Actualizar_BD(NSE_, "%", "SHARE UNIDADES", "0. Consolidado", "0. Cosmeticos", "UNIDADES (%)",
                            "MENSUAL", Periodo, valor_1, int.Parse(BD_Zoho.sCabecera48Meses[x].Substring(0, 4)));
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

    }
}
