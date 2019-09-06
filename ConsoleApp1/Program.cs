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
            DateTime HoraInicio, HoraFin;
            int xAño, xMes, xCiudad;
            string xMess;
            var obj = new BD_Zoho();
            var obj48 = new Output_48_Meses();

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

            //HoraInicio = DateTime.Now;
            //obj.Calcular_Periodos( xAño,xMes);
            //foreach (double item in obj.sShareValor)
            //{
            //    Console.WriteLine(item);
            //}
            //Console.WriteLine($"El proceso tardo: {DateTime.Now - HoraInicio}");
            
            Console.WriteLine("********** GENERACION DE ULTIMOS 48 MESES ************");
            xCiudad = 0;
            string[] xPeriodos48Meses = obj.Obtener_Ultimos_48_meses(xFecha[6].Year, xFecha[6].Month);
            
            Console.WriteLine($"{xPeriodos48Meses[0]}");
            Console.WriteLine($"{xPeriodos48Meses[1]}");

            //IDataReader x48Meses = obj48.Leer_Ultimos_48_Meses(xPeriodos48Meses[0], xPeriodos48Meses[1]);

            obj48.Leer_Ultimos_48_Meses(xPeriodos48Meses[0], xPeriodos48Meses[1]);
            double x = obj48.sdata48Meses_x_Region[0, 1];
            Console.WriteLine(x.ToString());
                        
            Console.ReadKey();
        }
    }
}
