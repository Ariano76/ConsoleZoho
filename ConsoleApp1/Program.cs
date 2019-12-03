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
        
        static void Main(string[] args)
        {
            int[] Codigo_Tipos_Importantes = { 158, 161, 215, 202, 237, 226 };
            
            DateTime HoraInicio;
            TimeSpan HoraFin;
            int xAño, xMes, xCiudad;
            string xMess;
            var obj = new BD_Zoho();
            var objHogar = new BL_HOGARES();
            var objUnidades = new BL_Unidades();
            var objHogar_1 = new BL_Hogares_V1();
            var obj48 = new Output_48_Meses();
            var objPPUDol = new BL_PPU_Dolares();

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
            xCiudad = 0;
            string[] xPeriodos48Meses = obj.Obtener_Ultimos_48_meses(xFecha[6].Year, xFecha[6].Month);
            
            Console.WriteLine($"{xPeriodos48Meses[0]}");

            Console.WriteLine($"{xPeriodos48Meses[1]}");

            //IDataReader x48Meses = obj48.Leer_Ultimos_48_Meses(xPeriodos48Meses[0], xPeriodos48Meses[1]);

            //*** INICIO - ESTE BLOQUE CORRESPONDE A CODIGO VALIDO PARA DESBLOQUEAR ***//
            obj.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            obj.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            obj.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            obj.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS

            //objHogar.Leer_Ultimos_48_Meses_Ciudad_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR NSE
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS HOGAR POR CATEGORIA
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y NSE
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, CATEGORIA Y MODALIDAD
            //                                                                                                     //*** FIN ***//
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y NSE
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_MODALIDAD(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y MODALIDAD            
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS_REGION(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL
            //objHogar.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CIUDAD, TIPOS Y TOTAL

            objUnidades.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE
            objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CATEGORIA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR CATEGORIA
            objUnidades.Leer_Ultimos_48_Meses_CIUDAD_CANAL_VENTA(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR MODALIDAD DE VENTA
            objUnidades.Leer_Ultimos_48_Meses_CIUDAD_TIPOS(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR TIPOS

            objPPUDol.Leer_Ultimos_48_Meses_CIUDAD_NSE(xPeriodos48Meses[0], xPeriodos48Meses[1]); // RESULTADOS POR NSE


            /////* LEYENDO LOS TIPOS ALMACENADOS EN UN ARREGLO*/
            ////string xTipos ="";
            ////foreach (int item in Codigo_Tipos_Importantes)
            ////{
            ////    xTipos += item + "," ;
            ////}
            ////Console.WriteLine($"Total Tipos Importantes:{xTipos}");



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
    }
}
