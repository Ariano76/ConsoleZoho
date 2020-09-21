using BL;
using BL.Periodos;
using System;
using System.Data.Common;

namespace ConsoleApp1
{
    class Program
    {
        public static DateTime HoraInicio;
        public static DateTime HoraStart;
        
        public static int Total_Record = 0;
        public static int Records = 0;

        static void Main(string[] args)
        {
            int[] Codigo_Tipos_Importantes = { 158, 161, 215, 202, 237, 226 };
            int[] Codigo_TIPOS = { 158, 161, 215, 202, 237, 226 }; // TIPOS 
            int[] Codigo_Categoria = { 123, 124, 127 }; // Categorias
            int Numero_Meses_YTD;
            string Ciudades_Codigos = "1,2,5"; // Ciudades
            string Ciudades_Cabecera = "[1],[2],[5]"; // Ciudades Cabecera

            TimeSpan HoraFin;
            int xAño, xMes;
            string xMess;

            var obj = new BD_Zoho();
            var objHogar = new BL_HOGARES();
            var objUnidades = new BL_Unidades();
            var obj48 = new Output_48_Meses();
            var objPPUDol = new BL_PPU_Dolares();
            var objPPUML = new BL_PPU_Moneda_Local();
            var objShareValor = new BL_Share_Dolares();
            var objShareUnidad = new BL_Share_Unidades();
            var objGastoMedioDol = new BL_Gasto_Medio_Dolares();
            var objGastoMedioML = new BL_Gasto_Medio_ML();
            var objUnidadesPromedioHogar = new BL_Unidades_Promedio_Hogar();
            var objPenetraciones = new BL_Penetracion_Hogar();
            var objMarcasShareValor = new BL_Share_Valor_Marca();
            var objMarcasShareUnidad = new BL_Share_Unidades_Marca();
            var objMarcasPPU_Valor = new BL_PPU_Moneda_Local_Marcas();
            var objMarcasPPU_Dolares = new BL_PPU_Dolares_Marcas();
            var objHogar_Marcas = new BL_Hogares_Marcas();
            var objUnidadPromedioHogar_Marca = new BL_Unidad_Promedio_Hogar_Marcas();
            var objGastoMedioDol_Marca = new BL_Gasto_Medio_DOL_Marcas();
            var objGastoMedioML_Marca = new BL_Gasto_Medio_ML_Marcas();
            var objPentraciones_Marca = new BL_Penetraciones_Marcas();
            var objPeridos_Share_Valor = new Share_Moneda_Local();
            var obj_Crear_BD_Peridos = new Poblar_Tabla_Factores();
            var objPeriodos_Marcas = new Marcas_Periodos();


            Console.WriteLine("Por favor Ingrese el Año de Proceso :");
            xAño = int.Parse(Console.ReadLine());           

            Console.WriteLine("Por favor Ingrese el Mes de Proceso :");
            xMes = int.Parse(Console.ReadLine());
            xMess = xMes.ToString();
            DateTime[] xFecha = obj.Restar_Meses_Fechas(xAño, xMes);
           
            obj.Periodo_Actual(xAño, xMes);
            Console.WriteLine(obj.sPeriodoActual[0]);            

            Console.WriteLine($"La fecha de inicio un año atras es: {xFecha[0].ToShortDateString()}\n" +
                $"La fecha de inicio dos años atras: {xFecha[1].ToShortDateString()}\n" +
                $"La fecha de inicio tres años atras: {xFecha[8].ToShortDateString()}\n" +
                $"La fecha de inicio tres meses atras: {xFecha[4].ToShortDateString()}\n" +
                $"La fecha de inicio tres meses atras año anterior: {xFecha[5].ToShortDateString()}\n" +
                $"La fecha de inicio seis meses atras : {xFecha[2].ToShortDateString()}\n" +
                $"La fecha de inicio seis meses atras año anterior: {xFecha[3].ToShortDateString()}\n" +
                $"La fecha de inicio seis meses atras Periodo 2 : {xFecha[14].ToShortDateString()}\n" +
                $"La fecha de Fin seis meses atras Periodo 2 : {xFecha[15].ToShortDateString()}\n" +                
                $"La fecha de inicio doce meses atras: {xFecha[9].ToShortDateString()}\n" +         
                $"La fecha de inicio doce meses atras dos años: {xFecha[10].ToShortDateString()}\n" +
                $"La fecha de inicio YTD : {xFecha[11].ToShortDateString()}\n" +
                $"La fecha de inicio YTD un año atras: {xFecha[12].ToShortDateString()}\n" +
                $"Periodo de inicio 48 meses atras (4 años): {xFecha[6].ToShortDateString()}\n" +
                $"Periodo de inicio 3 meses atras : {xFecha[7].ToShortDateString()}\n" +
                $"La fecha de Proceso Año Anterior es : {xFecha[0].Month} del año {xFecha[0].Year}");

            HoraInicio = DateTime.Now;
            //obj.Calcular_Periodos( xAño,xMes);
            //foreach (double item in obj.sShareValor)
            //{
            //    Console.WriteLine(item);
            //}
            Console.WriteLine($"Hora de Inicio: {HoraInicio}\n");

            Console.WriteLine("********** GENERACION DE ULTIMOS 48 MESES ************");

            string[] xPeriodos_Año_0 = obj.Obtener_periodos_año_0(xAño - 3);
            string[] xPeriodos_Año_1 = obj.Obtener_periodos_año_1(xAño - 2);
            string[] xPeriodos_Año_2 = obj.Obtener_periodos_año_2(xAño - 1);

            string[] xPeriodos_all_Años = obj.Obtener_periodos_all_años(xAño);

            string[] x3UltimosAños = obj.Obtener_3_Ultimos_años(xAño);
            string[] xPeriodos1Meses = obj.Obtener_Ultimos_1_mes(xAño, xMes);
            string[] xPeriodos1MesesAgo = obj.Obtener_Ultimos_1_mes_One_Year_Ago(xFecha[0].Year, xFecha[0].Month);
            string[] xPeriodos3Meses = obj.Obtener_Ultimos_3_meses(xFecha[7].Year, xFecha[7].Month);
            string[] xPeriodos3MesesAgo = obj.Obtener_Ultimos_3_meses_One_Year_Ago(xFecha[0].Year, xFecha[0].Month);
            string[] xPeriodos6Meses = obj.Obtener_Ultimos_6_meses(xFecha[2].Year, xFecha[2].Month);
            string[] xPeriodos6MesesAgo = obj.Obtener_Ultimos_6_meses_One_Year_Ago(xFecha[0].Year, xFecha[0].Month);
            string[] xPeriodos12Meses = obj.Obtener_Ultimos_12_meses(xFecha[9].Year, xFecha[9].Month);
            string[] xPeriodos12Meses_One_Year_Ago = obj.Obtener_Ultimos_12_meses_One_Year_Ago(xFecha[0].Year, xFecha[0].Month);
            string[] xPeriodosYTDMeses = obj.Obtener_YTD_meses(xFecha[11].Year, xFecha[11].Month);
            string[] xPeriodosYTDMesesAgo = obj.Obtener_YTD_meses_One_Year_Ago(xFecha[12].Year, xFecha[12].Month);

            string[] xPeriodos48Meses = obj.Obtener_Ultimos_48_meses(xFecha[6].Year, xFecha[6].Month);
            string[] xPeriodos_Inicio_Fin = obj.Obtener_Ultimos_48_meses_Factores_Capital(xFecha[6].Year, xFecha[6].Month, xAño, xMes);

            Console.WriteLine($"Periodos Año 0 : {xPeriodos_Año_0[0]}");
            Console.WriteLine($"Periodos Año 0 : {xPeriodos_Año_0[1]}\n");
            Console.WriteLine($"Periodos Año 1 : {xPeriodos_Año_1[0]}");
            Console.WriteLine($"Periodos Año 1 : {xPeriodos_Año_1[1]}\n");
            Console.WriteLine($"Periodos Año 2 : {xPeriodos_Año_2[0]}");
            Console.WriteLine($"Periodos Año 2 : {xPeriodos_Año_2[1]}\n");

            Console.WriteLine($"3 Ultimos Años : {x3UltimosAños[0]}");
            Console.WriteLine($"3 Ultimos Años : {x3UltimosAños[1]}\n");
            Console.WriteLine($"Periodos 1 Mes : {xPeriodos1Meses[0]}");
            Console.WriteLine($"Periodos 1 Mes : {xPeriodos1Meses[1]}\n");
            Console.WriteLine($"Periodos 1 Mes año anterior: {xPeriodos1MesesAgo[0]}");
            Console.WriteLine($"Periodos 1 Mes año anterior: {xPeriodos1MesesAgo[1]}\n");
            Console.WriteLine($"Periodos 3 Meses : {xPeriodos3Meses[0]}");
            Console.WriteLine($"Periodos 3 Meses : {xPeriodos3Meses[1]}\n");
            Console.WriteLine($"Periodos 3 Meses año anterior : {xPeriodos3MesesAgo[0]}");
            Console.WriteLine($"Periodos 3 Meses año anterior : {xPeriodos3MesesAgo[1]}\n");
            Console.WriteLine($"Periodos 6 Meses : {xPeriodos6Meses[0]}");
            Console.WriteLine($"Periodos 6 Meses : {xPeriodos6Meses[1]}\n");
            Console.WriteLine($"Periodos 6 Meses año anterior : {xPeriodos6MesesAgo[0]}");
            Console.WriteLine($"Periodos 6 Meses año anterior : {xPeriodos6MesesAgo[1]}\n");
            Console.WriteLine($"Periodos 12 Meses : {xPeriodos12Meses[0]}");
            Console.WriteLine($"Periodos 12 Meses : {xPeriodos12Meses[1]}\n");
            Console.WriteLine($"Periodos 12 Meses un año atras : {xPeriodos12Meses_One_Year_Ago[0]}");
            Console.WriteLine($"Periodos 12 Meses un año atras : {xPeriodos12Meses_One_Year_Ago[1]}\n");
            Console.WriteLine($"Periodos YTD Meses : {xPeriodosYTDMeses[0]}");
            Console.WriteLine($"Periodos YTD Meses : {xPeriodosYTDMeses[1]}\n");
            Console.WriteLine($"Periodos YTD Meses un año atras : {xPeriodosYTDMesesAgo[0]}");
            Console.WriteLine($"Periodos YTD Meses un año atras : {xPeriodosYTDMesesAgo[1]}\n");
            Console.WriteLine($"Periodos 48 Meses : {xPeriodos48Meses[0]}");            
            Console.WriteLine($"Periodos 48 Meses : {xPeriodos48Meses[1]}\n");
            Console.WriteLine($"Tabla Factor Cabecera : {xPeriodos_Inicio_Fin[0]}");
            Console.WriteLine($"Tabla Factor Periodos : {xPeriodos_Inicio_Fin[1]}\n");

            Numero_Meses_YTD = obj.Numero_Meses_en_YTD;
            Console.WriteLine($"Numero de meses en el YTD : {Numero_Meses_YTD}");
            //IDataReader x48Meses = obj48.Leer_Ultimos_48_Meses(xPeriodos48Meses[0], xPeriodos48Meses[1]);


            HoraStart = DateTime.Now;
            obj.Limpiar_BD_Resultados(); // RESULTADOS POR NSE
            Tiempo_Proceso("LIMPIANDO BASE DE DATOS ZOHO . .", HoraStart);
            Record_Progreso();


            //*** INICIO - ESTE BLOQUE CORRESPONDE A CODIGO VALIDO PARA DESBLOQUEAR ***//
            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            Tiempo_Proceso("DOLARES CIUDAD Y NSE . . .", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            Tiempo_Proceso("DOLARES CIUDAD Y CATEGORIA . . .", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            Tiempo_Proceso("DOLARES CIUDAD Y MODALIDAD . . .", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS
            Tiempo_Proceso("DOLARES CIUDAD Y TIPOS . . .", HoraStart);
            Record_Progreso();

            #region HOGARES            

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_Ciudad_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR NSE
            //Tiempo_Proceso("HOGARES CIUDAD Y NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR CATEGORIA
            //Tiempo_Proceso("HOGARES CIUDAD Y CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            //Tiempo_Proceso("HOGARES CIUDAD Y MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y NSE
            //Tiempo_Proceso("HOGARES CIUDAD, CATEGORIA Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("HOGARES CIUDAD, CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y NSE
            //Tiempo_Proceso("HOGARES CIUDAD, TIPO Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y MODALIDAD            
            //Tiempo_Proceso("HOGARES CIUDAD, TIPOS Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL
            //Tiempo_Proceso("HOGARES CIUDAD, TIPOS Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL
            //Tiempo_Proceso("HOGARES CIUDAD Y TIPOS . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region UNIDADES

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("UNIDADES CIUDAD Y NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            //Tiempo_Proceso("UNIDADES CIUDAD Y CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            //Tiempo_Proceso("UNIDADES CIUDAD Y MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS
            //Tiempo_Proceso("UNIDADES CIUDAD Y TIPOS . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU DOLARES

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (DOL) CIUDAD Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (DOL) CIUDAD Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y CATEGORIA
            //Tiempo_Proceso("PPU (DOL) NSE, CIUDAD Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (DOL) NSE Y CATEGORIA . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, REGION Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) CATEGORIA, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y REGION
            //Tiempo_Proceso("PPU (DOL) CATEGORIA Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA 
            //Tiempo_Proceso("PPU (DOL) CATEGORIA . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y TIPO
            //Tiempo_Proceso("PPU (DOL) NSE, CIUDAD Y TIPO .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPO
            //Tiempo_Proceso("PPU (DOL) NSE Y TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) TIPO, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION
            //Tiempo_Proceso("PPU (DOL) TIPO Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, MODALIDAD
            //Tiempo_Proceso("PPU (DOL) TIPO Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO
            //Tiempo_Proceso("PPU (DOL) TIPO . . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU SOLES            

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (ML) CIUDAD Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (ML) CIUDAD Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE, CIUDAD Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE Y CATEGORIA . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE, CIUDAD Y TIPO .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE Y TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO . . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE VALOR

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD Y NSE
            //Tiempo_Proceso("SHARE VALOR CIUDAD Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, MODALIDAD
            //Tiempo_Proceso("SHARE VALOR CIUDAD Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, CATEGORIA
            //Tiempo_Proceso("SHARE VALOR NSE, CIUDAD Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, TIPO
            //Tiempo_Proceso("SHARE VALOR NSE, CIUDAD Y TIPO .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("SHARE VALOR NSE Y CATEGORIA . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPO
            //Tiempo_Proceso("SHARE VALOR NSE Y TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE 
            //Tiempo_Proceso("SHARE VALOR NSE . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD 
            //Tiempo_Proceso("SHARE VALOR CIUDAD . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, REGION Y MODALIDAD
            //Tiempo_Proceso("SHARE VALOR CATEGORIA REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y REGION
            //Tiempo_Proceso("SHARE VALOR CATEGORIA Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("SHARE VALOR CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            //Tiempo_Proceso("SHARE VALOR CATEGORIA . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("SHARE VALOR TIPO, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION
            //Tiempo_Proceso("SHARE VALOR TIPO Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("SHARE VALOR TIPO Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y TOTAL
            //Tiempo_Proceso("SHARE VALOR TIPO Y TOTAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareValor.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD
            //Tiempo_Proceso("SHARE VALOR MODALIDAD . . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE UNIDADES

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD Y NSE
            //Tiempo_Proceso("SHARE UNIDAD CIUDAD Y NSE . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD CIUDAD Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, CATEGORIA
            //Tiempo_Proceso("SHARE UNIDAD NSE, CIUDAD Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, TIPO
            //Tiempo_Proceso("SHARE UNIDAD NSE, CIUDAD Y TIPO .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("SHARE UNIDAD NSE Y CATEGORIA . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPO
            //Tiempo_Proceso("SHARE UNIDAD NSE Y TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE 
            //Tiempo_Proceso("SHARE UNIDAD NSE . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD 
            //Tiempo_Proceso("SHARE UNIDAD CIUDAD . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, REGION Y MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD CATEGORIA REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y REGION
            //Tiempo_Proceso("SHARE UNIDAD CATEGORIA Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            //Tiempo_Proceso("SHARE UNIDAD CATEGORIA . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD TIPO, REGION Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION
            //Tiempo_Proceso("SHARE UNIDAD TIPO Y REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD TIPO Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y TOTAL
            //Tiempo_Proceso("SHARE UNIDAD TIPO Y TOTAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objShareUnidad.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD
            //Tiempo_Proceso("SHARE UNIDAD MODALIDAD . . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO DOLARES

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y CATEGORIA
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE, CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE, CIUDAD Y TIPOS", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE Y TIPOS .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("GASTO MEDIO DOLARES NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES TIPO, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_TIPO_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION 
            //Tiempo_Proceso("GASTO MEDIO DOLARES TIPO Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES TIPO Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO 
            //Tiempo_Proceso("GASTO MEDIO DOLARES TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES CIUDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, CIUDAD Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES CATEGORIA, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y CIUDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES CATEGORIA Y CIUDAD . ", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES CATEGORIA Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_CATEGORIAS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA 
            //Tiempo_Proceso("GASTO MEDIO DOLARES CATEGORIAS . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD Y CIUDAD
            //Tiempo_Proceso("GASTO MEDIO DOLARES MODALIDAD Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD 
            //Tiempo_Proceso("GASTO MEDIO DOLARES MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO MONEDA LOCAL

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y CATEGORIA
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE, CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE, CIUDAD Y TIPOS", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE Y TIPOS .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL TIPO, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_TIPO_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION 
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL TIPO Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL TIPO Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO 
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL CIUDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, CIUDAD Y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL CATEGORIA, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y CIUDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL CATEGORIA Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y MODALIDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL CATEGORIA Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_CATEGORIAS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA 
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL CATEGORIAS . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD Y CIUDAD
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL MODALIDAD Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD 
            //Tiempo_Proceso("GASTO MEDIO MONEDALOCAL MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region UNIDADES PROMEDIO HOGAR

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y CATEGORIA
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE, CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y TIPOS
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE, CIUDAD Y TIPOS", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE Y TIPOS .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR TIPO, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_TIPO_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION 
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR TIPO Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR TIPO Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO 
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR CIUDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, CIUDAD Y MODALIDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR CATEGORIA, CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y CIUDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR CATEGORIA Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y MODALIDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR CATEGORIA Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_CATEGORIAS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA 
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR CATEGORIAS . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD Y CIUDAD
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR MODALIDAD Y CIUDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadesPromedioHogar.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD 
            //Tiempo_Proceso("UNIDADES PROMEDIO HOGAR MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PENETRACIONES HOGAR

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1], xPeriodos_Inicio_Fin[0], xPeriodos_Inicio_Fin[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR NSE . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE_CATEGORIA_CONSOLIDADO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR NSE Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE_TIPO_CONSOLIDADO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR NSE Y TIPO .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPOS
            //Tiempo_Proceso("PENETRACION HOGAR NSE, CIUDAD Y CATEGORIA .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD Y TIPOS
            //Tiempo_Proceso("PENETRACION HOGAR NSE, CIUDAD Y TIPOS .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_NSE_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1], xPeriodos_Inicio_Fin[0], xPeriodos_Inicio_Fin[1]); // RESULTADOS POR NSE Y CIUDAD 
            //Tiempo_Proceso("PENETRACION HOGAR NSE Y CIUDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD Y CIUDAD
            //Tiempo_Proceso("PENETRACION HOGAR MODALIDAD Y CIUDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_MODALIDAD_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, CIUDAD Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR CATEGORIA, CIUDAD Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TIPO_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR TIPO, CIUDAD Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_MODALIDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR CATEGORIA Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR TIPO Y MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD 
            //Tiempo_Proceso("PENETRACION HOGAR MODALIDAD . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO 
            //Tiempo_Proceso("PENETRACION HOGAR TIPO, REGION . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TIPO_MODALIDAD_CONSOLIDADO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR TIPO Y MODALIDAD .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TIPO_CONSOLIDADO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR CATEGORIA, CIUDAD, MODALIDAD . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_CATEGORIA_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR CATEGORIA, CIUDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION HOGAR CATEGORIA, MODALIDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_CATEGORIA_CONSOLIDADO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1], xPeriodos_Inicio_Fin[0], xPeriodos_Inicio_Fin[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR CIUDAD . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPenetraciones.Leer_Ultimos_48_Meses_TOTAL_PAIS(xPeriodos48Meses[0], xPeriodos48Meses[1], xPeriodos_Inicio_Fin[0], xPeriodos_Inicio_Fin[1]); // RESULTADOS POR TIPO Y MODALIDAD
            //Tiempo_Proceso("PENETRACION CONSOLIDADO HOGAR PAIS . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS VALOR TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1]); // 
            //objMarcasShareValor.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE VALORES MARCA TOTAL COSMETICOS . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS UNIDADES TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1]); // 
            //objMarcasShareUnidad.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE UNIDADES MARCA TOTAL COSMETICOS . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS VALOR TOTAL CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_TIPO_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA TIPO CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA CATEGORIA CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Capital(xPeriodos3Meses[1]); // 
            //objMarcasShareValor.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE VALORES MARCA TOTAL COSMETICOS CAPITAL .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS UNIDADES TOTAL CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_TIPO_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA TIPO CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA CATEGORIA CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Capital(xPeriodos3Meses[1]); // 
            //objMarcasShareUnidad.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE UNIDADES MARCA TOTAL COSMETICOS CAPITAL .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS VALOR TOTAL CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_TIPO_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA TIPO CIUDADES . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item); // 
            //    objMarcasShareValor.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE VALORES MARCA CATEGORIA CIUDADES . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareValor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Region(xPeriodos3Meses[1]); // 
            //objMarcasShareValor.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE VALORES MARCA TOTAL COSMETICOS CIUDADES .", HoraStart);
            //Record_Progreso();

            #endregion

            #region SHARE MARCAS UNIDADES TOTAL CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_TIPO_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA TIPO CIUDADES . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item); // 
            //    objMarcasShareUnidad.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("SHARE UNIDADES MARCA CATEGORIA CIUDADES . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasShareUnidad.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Region(xPeriodos3Meses[1]); // 
            //objMarcasShareUnidad.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("SHARE UNIDADES MARCA TOTAL COSMETICOS CIUDADES .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU MONEDA LOCAL MARCAS VALOR TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Valor.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Valor.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA CATEGORIA . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1], 1); // 
            //objMarcasPPU_Valor.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], 1);
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TOTAL COSMETICOS .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU MONEDA LOCAL MARCAS VALOR TOTAL CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_TIPO_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TIPO CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_CATEGORIA_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA CATEGORIA CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Capital(xPeriodos3Meses[1], 1); // 
            //objMarcasPPU_Valor.Leer_Ultimos_48_Meses_TOTAL_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], 1);
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TOTAL CAPITAL .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU MONEDA LOCAL MARCAS VALOR TOTAL CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_TIPO_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TIPO REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item, 1); // 
            //    objMarcasPPU_Valor.Leer_Ultimos_48_Meses_CATEGORIA_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 1);
            //}
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA CATEGORIA REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Valor.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Region(xPeriodos3Meses[1], 1); // 
            //objMarcasPPU_Valor.Leer_Ultimos_48_Meses_TOTAL_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], 1);
            //Tiempo_Proceso("PPU VALORES LOCAL MARCA TOTAL REGION .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU DOLARES MARCAS VALOR TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA TIPO . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA CATEGORIA . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1], 2); // 
            //objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], 2);
            //Tiempo_Proceso("PPU DOLARES MARCA TOTAL COSMETICOS . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU DOLARES MARCAS VALOR TOTAL CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_TIPO_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA TIPO CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_CATEGORIA_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA CATEGORIA CAPITAL . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Capital(xPeriodos3Meses[1], 2); // 
            //objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_TOTAL_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], 2);
            //Tiempo_Proceso("PPU DOLARES MARCA TOTAL CAPITAL . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PPU DOLARES MARCAS VALOR TOTAL CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_TIPO_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA TIPO REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item, 2); // 
            //    objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_CATEGORIA_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, 2);
            //}
            //Tiempo_Proceso("PPU DOLARES MARCA CATEGORIA REGION . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objMarcasPPU_Dolares.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Region(xPeriodos3Meses[1], 2); // 
            //objMarcasPPU_Dolares.Leer_Ultimos_48_Meses_TOTAL_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], 2);
            //Tiempo_Proceso("PPU DOLARES MARCA TOTAL REGION . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region HOGARES MARCAS TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA TIPO . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA CATEGORIAS . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1]); // 
            //objHogar_Marcas.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("HOGARES MARCA TOTAL COSMETICOS . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region HOGARES MARCAS TOTAL CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item); //  
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_TIPO_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA TIPO CAPITAL . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_CATEGORIA_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA CATEGORIAS CAPITAL . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Capital(xPeriodos3Meses[1]); // 
            //objHogar_Marcas.Leer_Ultimos_48_Meses_TOTAL_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("HOGARES MARCA TOTAL COSMETICOS CAPITAL . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region HOGARES MARCAS TOTAL CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA TIPO CIUDADES . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item); // 
            //    objHogar_Marcas.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("HOGARES MARCA CATEGORIAS CIUDADES . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar_Marcas.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos_Region(xPeriodos3Meses[1]); // 
            //objHogar_Marcas.Leer_Ultimos_48_Meses_TOTAL_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("HOGARES MARCA TOTAL COSMETICOS CIUDADES . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region UNIDAD PROMEDIO POR HOGAR - MARCAS TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA CATEGORIAS . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1]); // 
            //objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TOTAL COSMETICOS .", HoraStart);
            //Record_Progreso();

            #endregion

            #region UNIDAD PROMEDIO POR HOGAR - MARCAS CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_TIPO_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TIPO CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_Categoria_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA CATEGORIAS CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Capital(xPeriodos3Meses[1]); // 
            //objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TOTAL COSMETICOS CAPITAL", HoraStart);
            //Record_Progreso();

            #endregion

            #region UNIDAD PROMEDIO POR HOGAR - MARCAS CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_TIPO_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TIPO REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item); // 
            //    objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_Categoria_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA CATEGORIAS REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidadPromedioHogar_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Region(xPeriodos3Meses[1]); // 
            //objUnidadPromedioHogar_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Region(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("UNIDAD PROMEDIO HOGAR MARCA TOTAL COSMETICOS REGION", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO DOLARES - MARCAS TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item, "2"); // 
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item, "2"); // 
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA CATEGORIAS . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1], "2"); // 
            //objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], "2");
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TOTAL COSMETICOS .", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO DOLARES - MARCAS CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item, "2"); // 
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_TIPO_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TIPO CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item, "2"); // 
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_Categoria_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA CATEGORIAS CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Capital(xPeriodos3Meses[1], "2"); // 
            //objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], "2");
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TOTAL COSMETICOS CAPITAL", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO DOLARES - MARCAS CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item, "2"); //  
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_TIPO_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TIPO REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item, "2"); // 
            //    objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_Categoria_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "2");
            //}
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA CATEGORIAS REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioDol_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Region(xPeriodos3Meses[1], "2"); // 
            //objGastoMedioDol_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], "2");
            //Tiempo_Proceso("GASTO MEDIO DOLARES MARCA TOTAL COSMETICOS REGION", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO MONEDA LOCAL - MARCAS TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item, "1"); // 
            //    objGastoMedioML_Marca.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objGastoMedioML_Marca.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objGastoMedioML_Marca.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TIPO . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item, "1"); // 
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA CATEGORIAS . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1], "1"); // 
            //objGastoMedioML_Marca.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], "1");
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TOTAL COSMETICOS .", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO MONEDA LOCAL - MARCAS CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item, "1"); //  
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_TIPO_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TIPO CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item, "1"); // 
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_Categoria_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA CATEGORIAS CAPITAL .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Capital(xPeriodos3Meses[1], "1"); // 
            //objGastoMedioML_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Capital(xPeriodos48Meses[0], xPeriodos48Meses[1], "1");
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TOTAL COSMETICOS CAPITAL", HoraStart);
            //Record_Progreso();

            #endregion

            #region GASTO MEDIO MONEDA LOCAL - MARCAS CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item, "1"); // 
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_TIPO_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TIPO REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item, "1"); // 
            //    objGastoMedioML_Marca.Leer_Ultimos_48_Meses_Categoria_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], item, "1");
            //}
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA CATEGORIAS REGION .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objGastoMedioML_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Region(xPeriodos3Meses[1], "1"); // 
            //objGastoMedioML_Marca.Leer_Ultimos_48_Meses_Cosmeticos_Region(xPeriodos48Meses[0], xPeriodos48Meses[1], "1");
            //Tiempo_Proceso("GASTO MEDIO SOLES MARCA TOTAL COSMETICOS REGION", HoraStart);
            //Record_Progreso();

            #endregion

            #region PENETRACIONES - MARCAS TOTAL PAIS

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Recuperar_Marcas_Grupo_Belcorp(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Recuperar_Marcas_Grupo_Lauder(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Recuperar_Marcas_Grupo_Loreal(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_TIPO_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA TIPO . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_CATEGORIA_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA CATEGORIAS . . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(xPeriodos3Meses[1]); // 
            //objPentraciones_Marca.Leer_Ultimos_48_Meses_COSMETICOS_TOTAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("PENETRACIONES MARCA TOTAL . . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PENETRACIONES - MARCAS CAPITAL

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Capital(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_TIPO_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA TIPO CAPITAL . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Capital(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_CATEGORIA_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA CATEGORIAS CAPITAL . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Capital(xPeriodos3Meses[1]); // 
            //objPentraciones_Marca.Leer_Ultimos_48_Meses_COSMETICOS_CAPITAL(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("PENETRACIONES MARCA TOTAL CAPITAL . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region PENETRACIONES - MARCAS CIUDADES

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_TIPOS)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Tipo_Region(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA TIPO REGION . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //foreach (var item in Codigo_Categoria)
            //{
            //    objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Categoria_Region(xPeriodos3Meses[1], item); // 
            //    objPentraciones_Marca.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1], item);
            //}
            //Tiempo_Proceso("PENETRACIONES MARCA CATEGORIAS REGION . . .", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPentraciones_Marca.Recuperar_Marcas_Top_5_Retail_x_Total_Region(xPeriodos3Meses[1]); // 
            //objPentraciones_Marca.Leer_Ultimos_48_Meses_COSMETICOS_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            //Tiempo_Proceso("PENETRACIONES MARCA TOTAL REGION . . .", HoraStart);
            //Record_Progreso();

            #endregion

            #region CREA TABLA RESULTADOS POR PERIODOS HOGARES
            // CREACION DE FACTORES PARA LOS DIFERENTES PERIODOS 

            //HoraStart = DateTime.Now;
            //obj_Crear_BD_Peridos.Crear_Tabla_Datos_Periodos(xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], xPeriodos_Año_0[1], xPeriodos_Año_1[1], xPeriodos_Año_2[1]);

            //obj_Crear_BD_Peridos.Crear_Tabla_Factores(Ciudades_Codigos, Ciudades_Cabecera, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], xPeriodos_Año_0[1], xPeriodos_Año_1[1], xPeriodos_Año_2[1]);

            //Tiempo_Proceso("CREACION TABLA RESULTADOS PERIODOS HOGARES . .", HoraStart);
            //Record_Progreso();
            #endregion

            #region PERIODOS

            //HoraStart = DateTime.Now;
            //objPeridos_Share_Valor.Recuperar_Codigos_NSE();
            //objPeridos_Share_Valor.Recuperar_Codigos_Categoria();
            //objPeridos_Share_Valor.Recuperar_Codigos_Modalidad();
            //objPeridos_Share_Valor.Recuperar_Codigos_Tipos();
            //objPeridos_Share_Valor.Universo_Periodos();

            //// HOGAR PAIS
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Hogares("1,2,5");
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Hogares("1");
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Hogares("2,5");

            //Tiempo_Proceso("PERIODOS HOGARES . . . . .", HoraStart);
            //Record_Progreso();

            //// UNIDADES PAIS
            //HoraStart = DateTime.Now;
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Unidades("1,2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1],Numero_Meses_YTD);
            //// MONEDA LOCAL            
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("1,2,5", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            ////DOLARES
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("1,2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);

            //Tiempo_Proceso("PERIODOS DATOS VALOR PAIS . . .", HoraStart);
            //Record_Progreso();

            //// UNIDADES CAPITAL
            //HoraStart = DateTime.Now;
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Unidades("1", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1],Numero_Meses_YTD);
            //// MONEDA LOCAL
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("1", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            ////DOLARES
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("1", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);

            //Tiempo_Proceso("PERIODOS DATOS VALOR CAPITAL . . .", HoraStart);
            //Record_Progreso();

            //// UNIDADES CIUDADES
            //HoraStart = DateTime.Now;
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Unidades("2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            //// MONEDA LOCAL
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("2,5", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            ////DOLARES
            //objPeridos_Share_Valor.Periodos_Cosmeticos_Total_Valores("2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);

            //Tiempo_Proceso("PERIODOS DATOS VALOR CIUDADES . . .", HoraStart);
            //Record_Progreso();

            #endregion


            #region PERIODOS POR MARCAS

            HoraStart = DateTime.Now;
            objPeriodos_Marcas.Recuperar_Codigos_NSE();
            objPeriodos_Marcas.Recuperar_Codigos_Categoria();
            //objPeriodos_Marcas.Recuperar_Codigos_Modalidad();
            objPeriodos_Marcas.Recuperar_Codigos_Tipos();
            //objPeriodos_Marcas.Universo_Periodos();
            //objPeriodos_Marcas.Recuperar_Marcas_Top_5_Retail_x_Total_Cosmeticos(1);
            objPeriodos_Marcas.Recuperar_Marcas_Grupo_Belcorp(); // 
            objPeriodos_Marcas.Recuperar_Marcas_Grupo_Lauder(); // 
            objPeriodos_Marcas.Recuperar_Marcas_Grupo_Loreal(); // 


            //// HOGAR PAIS
            //objPeriodos_Marcas.Periodos_Cosmeticos_Total_Hogares("1,2,5");
            //objPeriodos_Marcas.Periodos_Cosmeticos_Total_Hogares("1");
            //objPeriodos_Marcas.Periodos_Cosmeticos_Total_Hogares("2,5");
            objPeriodos_Marcas.Hogares("1,2,5");
            objPeriodos_Marcas.Hogares("1");
            objPeriodos_Marcas.Hogares("2,5");

            Tiempo_Proceso("PERIODOS HOGARES . . . . .", HoraStart);
            Record_Progreso();

            //// UNIDADES PAIS
            HoraStart = DateTime.Now;
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Unidades("1,2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            //// MONEDA LOCAL            
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Valores("1,2,5", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            Tiempo_Proceso("PERIODOS DATOS MARCAS PAIS . . .", HoraStart);
            Record_Progreso();

            //// UNIDADES CAPITAL
            HoraStart = DateTime.Now;
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Unidades("1", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            //// MONEDA LOCAL
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Valores("1", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            Tiempo_Proceso("PERIODOS DATOS MARCAS CAPITAL . . .", HoraStart);
            Record_Progreso();

            //// UNIDADES CIUDADES
            HoraStart = DateTime.Now;
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Unidades("2,5", x3UltimosAños[0], x3UltimosAños[1], 2, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);
            //// MONEDA LOCAL
            objPeriodos_Marcas.Periodos_Cosmeticos_Total_Valores("2,5", x3UltimosAños[0], x3UltimosAños[1], 1, xPeriodos12Meses_One_Year_Ago[1], xPeriodos12Meses[1], xPeriodos6MesesAgo[1], xPeriodos6Meses[1], xPeriodos3MesesAgo[1], xPeriodos3Meses[1], xPeriodos1MesesAgo[1], xPeriodos1Meses[1], xPeriodosYTDMesesAgo[1], xPeriodosYTDMeses[1], Numero_Meses_YTD);

            Tiempo_Proceso("PERIODOS DATOS MARCAS CIUDADES . . .", HoraStart);
            Record_Progreso();

            #endregion











            /////* LEYENDO LOS TIPOS ALMACENADOS EN UN ARREGLO*/
            ////string xTipos ="";
            ////foreach (int item in Codigo_Tipos_Importantes)
            ////{
            ////    xTipos += item + "," ;
            ////}
            ////Console.WriteLine($"Total Tipos Importantes:{xTipos}");
            ///
            Console.WriteLine($"***********************\n");
            Console.WriteLine(obj48.sdata48Meses_x_Region.GetUpperBound(1) - obj.sdata48Meses_x_Region_NSE.GetLowerBound(1) + 1);
            Console.WriteLine($"Numero de filas del Array: {obj.sdata48Meses_x_Region_NSE.GetLength(0)}");
            Console.WriteLine($"Numero de columnas del Array: {obj.sdata48Meses_x_Region_NSE.GetLength(1)}");
            Console.WriteLine($"Totales por NSE \n");

            //for (int i = 0; i < obj.sdata48Meses_x_Region_NSE.GetLength(1); i++)
            //{
            //    Console.WriteLine($"{obj.sdata48Meses_x_Region_NSE[0, i]} - {obj.sdata48Meses_x_Region_NSE[5, i]} - ({obj.sdata48Meses_x_Region_NSE[i, i] + obj.sdata48Meses_x_Region_NSE[5, i]})");
            //}
            //Console.WriteLine($"Totales por CANAL DE VENTA \n");
            //for (int i = 0; i < obj.sdata48Meses_x_Region_Modalidad_Mes.GetLength(1); i++)
            //{
            //    Console.WriteLine($"{obj.sdata48Meses_x_Region_Modalidad_Mes[0, i]} - {obj.sdata48Meses_x_Region_Modalidad_Mes[1, i]} - {obj.sdata48Meses_x_Region_Modalidad_Mes[2, i]}");
            //}
            //Console.WriteLine($"Totales por TIPOS \n");
            //for (int i = 0; i < obj.sdata48Meses_x_Region_Tipos_Mes.GetLength(1); i++)
            //{
            //    Console.WriteLine($"{obj.sdata48Meses_x_Region_Tipos_Mes[0, i]} - {obj.sdata48Meses_x_Region_Tipos_Mes[1, i]} - {obj.sdata48Meses_x_Region_Tipos_Mes[2, i]}");
            //}

            Console.WriteLine(obj.resultadoBD);
            HoraFin = DateTime.Now - HoraInicio;
            Console.WriteLine($"El proceso tardo: {HoraFin}");
            Console.WriteLine($"El proceso tardo: {HoraFin.TotalSeconds}");
            Console.ReadKey();
        }

        static void Tiempo_Proceso(string Proceso, DateTime Inicio)
        {
            TimeSpan HoraFin;
            HoraFin = DateTime.Now - Inicio;
            Console.Write($"El Proceso {Proceso} tardo {HoraFin} segundos.");
        }
        static void Record_Progreso()
        {
            var objGenerico = new Procesos_Genericos();

            Records = objGenerico.Leer_Total_Registros_BD();
            Console.WriteLine($" - Total Registros : {Records - Total_Record}");
            Total_Record =+ Records;
        }
    }
}
