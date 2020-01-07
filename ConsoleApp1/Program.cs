using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using BL;

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
            TimeSpan HoraFin;
            int xAño, xMes;
            string xMess;

            var obj = new BD_Zoho();
            var objHogar = new BL_HOGARES();
            var objUnidades = new BL_Unidades();
            var objHogar_1 = new BL_Hogares_V1();
            var obj48 = new Output_48_Meses();
            var objPPUDol = new BL_PPU_Dolares();
            var objPPUML = new BL_PPU_Moneda_Local();
            var objShareValor = new BL_Share_Dolares();


            Console.WriteLine("Por favor Ingrese el Año de Proceso :");
            xAño = int.Parse(Console.ReadLine());
            Console.WriteLine("Por favor Ingrese el Mes de Proceso :");
            xMes = int.Parse(Console.ReadLine());
            xMess = xMes.ToString();
            DateTime[] xFecha = obj.Restar_Meses_Fechas(xAño, xMes);

            obj.Periodo_Actual(xAño, xMes);
            Console.WriteLine(obj.sPeriodoActual[0]);

            Console.WriteLine($"La fecha de un año atras es: {xFecha[0].ToShortDateString()}\n" +
                $"La fecha dos años atras: {xFecha[1].ToShortDateString()}\n" +
                $"La fecha seis meses atras: {xFecha[2].ToShortDateString()}\n" +
                $"La fecha seis meses atras dos años: {xFecha[3].ToShortDateString()}\n" +
                $"La fecha tres meses atras: {xFecha[4].ToShortDateString()}\n" +
                $"La fecha tres meses atras dos años: {xFecha[5].ToShortDateString()}\n" +
                $"Periodo de inicio 48 meses atras (4 años): {xFecha[6].ToShortDateString()}\n" +
                $"El mes es {xFecha[0].Month} del año {xFecha[0].Year}");

            HoraInicio = DateTime.Now;
            //obj.Calcular_Periodos( xAño,xMes);
            //foreach (double item in obj.sShareValor)
            //{
            //    Console.WriteLine(item);
            //}
            Console.WriteLine($"Hora de Inicio: {HoraInicio}");

            Console.WriteLine("********** GENERACION DE ULTIMOS 48 MESES ************");

            string[] xPeriodos48Meses = obj.Obtener_Ultimos_48_meses(xFecha[6].Year, xFecha[6].Month);

            Console.WriteLine($"{xPeriodos48Meses[0]}");

            Console.WriteLine($"{xPeriodos48Meses[1]}\n");

            //IDataReader x48Meses = obj48.Leer_Ultimos_48_Meses(xPeriodos48Meses[0], xPeriodos48Meses[1]);

            //*** INICIO - ESTE BLOQUE CORRESPONDE A CODIGO VALIDO PARA DESBLOQUEAR ***//
            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            Tiempo_Proceso("DOLARES CIUDAD Y NSE", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            Tiempo_Proceso("DOLARES CIUDAD Y CATEGORIA", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            Tiempo_Proceso("DOLARES CIUDAD Y MODALIDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            obj.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS
            Tiempo_Proceso("DOLARES CIUDAD Y TIPOS", HoraStart);
            Record_Progreso();

            //** RESULTADOS HOGARES **//

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_Ciudad_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR NSE
            //Tiempo_Proceso("HOGARES CIUDAD Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR CATEGORIA
            //Tiempo_Proceso("HOGARES CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            //Tiempo_Proceso("HOGARES CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y NSE
            //Tiempo_Proceso("HOGARES CIUDAD, CATEGORIA Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("HOGARES CIUDAD, CATEGORIA Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y NSE
            //Tiempo_Proceso("HOGARES CIUDAD, TIPO Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y MODALIDAD            
            //Tiempo_Proceso("HOGARES CIUDAD, TIPOS Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL
            //Tiempo_Proceso("HOGARES CIUDAD, TIPOS Y REGION", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL
            //Tiempo_Proceso("HOGARES CIUDAD Y TIPOS", HoraStart);
            //Record_Progreso();

            ////** UNIDADES **//
            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("UNIDADES CIUDAD Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            //Tiempo_Proceso("UNIDADES CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            //Tiempo_Proceso("UNIDADES CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objUnidades.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS
            //Tiempo_Proceso("UNIDADES CIUDAD Y TIPOS", HoraStart);
            //Record_Progreso();

            ////****PPU DOLARES**** //

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (DOL) CIUDAD Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (DOL) CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y CATEGORIA
            //Tiempo_Proceso("PPU (DOL) NSE, CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (DOL) NSE Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, REGION Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) CATEGORIA, REGION Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) CATEGORIA Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y REGION
            //Tiempo_Proceso("PPU (DOL) CATEGORIA Y REGION", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA 
            //Tiempo_Proceso("PPU (DOL) CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y TIPO
            //Tiempo_Proceso("PPU (DOL) NSE, CIUDAD Y TIPO", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPO
            //Tiempo_Proceso("PPU (DOL) NSE Y TIPO", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            //Tiempo_Proceso("PPU (DOL) TIPO, REGION Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION
            //Tiempo_Proceso("PPU (DOL) TIPO Y REGION", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, MODALIDAD
            //Tiempo_Proceso("PPU (DOL) TIPO Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUDol.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO
            //Tiempo_Proceso("PPU (DOL) TIPO", HoraStart);
            //Record_Progreso();

            //// **** PPU SOLES **** //

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (ML) CIUDAD Y NSE", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            //Tiempo_Proceso("PPU (ML) CIUDAD Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, REGION Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE, CIUDAD Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE Y CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA, REGION Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA Y REGION", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) CATEGORIA", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE, CIUDAD Y TIPO", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) NSE Y TIPO", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO, REGION Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO Y REGION", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO Y MODALIDAD", HoraStart);
            //Record_Progreso();

            //HoraStart = DateTime.Now;
            //objPPUML.Leer_Ultimos_48_Meses_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            //Tiempo_Proceso("PPU (ML) TIPO", HoraStart);
            //Record_Progreso();

            //****SHARE VALOR **** //

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD Y NSE
            Tiempo_Proceso("SHARE VALOR CIUDAD Y NSE", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CIUDAD_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, MODALIDAD
            Tiempo_Proceso("SHARE VALOR CIUDAD Y MODALIDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_NSE_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, CATEGORIA
            Tiempo_Proceso("SHARE VALOR NSE, CIUDAD Y CATEGORIA", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_NSE_CIUDAD_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE, CIUDAD, TIPO
            Tiempo_Proceso("SHARE VALOR NSE, CIUDAD Y TIPO", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_NSE_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y CATEGORIA
            Tiempo_Proceso("SHARE VALOR NSE Y CATEGORIA", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_NSE_TIPO(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE Y TIPO
            Tiempo_Proceso("SHARE VALOR NSE Y TIPO", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE 
            Tiempo_Proceso("SHARE VALOR NSE", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CIUDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD 
            Tiempo_Proceso("SHARE VALOR CIUDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA, REGION Y MODALIDAD
            Tiempo_Proceso("SHARE VALOR CATEGORIA REGION Y MODALIDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y REGION
            Tiempo_Proceso("SHARE VALOR CATEGORIA Y REGION ", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA Y MODALIDAD
            Tiempo_Proceso("SHARE VALOR CATEGORIA Y MODALIDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            Tiempo_Proceso("SHARE VALOR CATEGORIA", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_TIPO_REGION_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO, REGION Y MODALIDAD
            Tiempo_Proceso("SHARE VALOR TIPO, REGION Y MODALIDAD", HoraStart);
            Record_Progreso();

            HoraStart = DateTime.Now;
            objShareValor.Leer_Ultimos_48_Meses_TIPO_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPO Y REGION
            Tiempo_Proceso("SHARE VALOR TIPO Y REGION", HoraStart);
            Record_Progreso();









            /////* LEYENDO LOS TIPOS ALMACENADOS EN UN ARREGLO*/
            ////string xTipos ="";
            ////foreach (int item in Codigo_Tipos_Importantes)
            ////{
            ////    xTipos += item + "," ;
            ////}
            ////Console.WriteLine($"Total Tipos Importantes:{xTipos}");
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
