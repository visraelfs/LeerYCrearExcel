using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReporteMontosDevengar.Clases;

namespace ReporteMontosDevengar
{
    class Program
    {
        static void Main(string[] args)
        {
            string rutaArchivoSalida = @"C:\Users\USUARIO\Desktop\Montos Por Devengar Completo.xlsx";
            string rutaArchivoOrigen = @"C:\Users\USUARIO\Desktop\Montos por Devengar por EP.xlsx";
            try
            {

                GenerarReporteMontoDevengar reporte = new GenerarReporteMontoDevengar("Montos por Devengar por EP", true, rutaArchivoOrigen, rutaArchivoSalida);

                reporte.generarInforme();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Console.WriteLine("Presione una tecla para continuar...");
                //Console.ReadKey();

                ProcessStartInfo startInfo = new ProcessStartInfo();                
                Process.Start(rutaArchivoSalida);
            }
        }
    }
}
