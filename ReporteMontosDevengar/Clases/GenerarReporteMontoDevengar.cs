using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReporteMontosDevengar.Interfaces;
using ExcelDataReader;
using ReporteMontosDevengar.Models;
using ClosedXML.Excel;

namespace ReporteMontosDevengar.Clases
{
    public class GenerarReporteMontoDevengar : ILectorExcel
    {

        private string hojaInforme = string.Empty;
        private bool useHeaderRow = false;
        private string rutaArchivoSalida = string.Empty;
        private string rutaArchivoOrigen = string.Empty;


        public GenerarReporteMontoDevengar(string hojaInforme, bool useHeaderRow,string rutaArchivoOrigen, string rutaArchivoSalida)
        {
            this.hojaInforme = hojaInforme;
            this.useHeaderRow = useHeaderRow;
            this.rutaArchivoSalida = rutaArchivoSalida;
            this.rutaArchivoOrigen = rutaArchivoOrigen;
        }

        public void generarInforme()
        {
            try
            {
                string rutaArchivo = this.rutaArchivoOrigen;
                MemoryStream ms = new MemoryStream();

                FileStream file = new FileStream(rutaArchivo, FileMode.Open, System.IO.FileAccess.Read);

                file.CopyTo(ms);


                DataSet ds = leerExcel(ms);


                int max = ds.Tables[this.hojaInforme].Rows.Cast<DataRow>().Max(x => Int32.Parse(x["Meses por Devengar"].ToString()));
                int renglones = ds.Tables[this.hojaInforme].Rows.Count + 1;


                var workbook = new XLWorkbook(rutaArchivo);
                var hoja1 = workbook.Worksheet(1);

                int mesActual = DateTime.Now.Month;
                int contadorMes = mesActual;
                int anioActual = DateTime.Now.Year;
                int celdaCabecera = 10;

                string formatoMoneda = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \" - \"??_-;_-@_-";


                #region Establece las cabeceras de las fechas, dependiendo del plazo maximo que tenga el archivo
                for (int i = 0; i < max; i++)
                {
                    var format = "mmmm-yyyy";
                    hoja1.Cell(1, celdaCabecera).Value = obtenerMes(contadorMes) + " " + anioActual;
                    hoja1.Cell(1, celdaCabecera).Style.NumberFormat.Format = format;
                    hoja1.Cell(1, celdaCabecera).Style.Font = (hoja1.Cell(1, 1).Style.Font);
                    hoja1.Cell(1, celdaCabecera).Style.Fill.BackgroundColor = (workbook.Worksheet(1).Cell(1, 1).Style.Fill.BackgroundColor);


                    contadorMes++;
                    if (contadorMes > 12)
                    {
                        anioActual++;
                        contadorMes = 1;
                    }
                    celdaCabecera++;
                }
                #endregion

                #region Establecer valores por mes

                for (int renglon = 2; renglon <= renglones; renglon++)
                {
                    decimal valorMensual = Decimal.Parse(hoja1.Cell(renglon, 10).Value.ToString()) / Decimal.Parse(hoja1.Cell(renglon, 9).Value.ToString());
                    for (int columna = 11; columna < (Int32.Parse(hoja1.Cell(renglon, 9).Value.ToString()) + 11); columna++)
                    {
                        hoja1.Cell(renglon, columna).Value = valorMensual;
                        hoja1.Cell(renglon, columna).Style.NumberFormat.Format = formatoMoneda;
                        hoja1.Cell(renglon, columna).Style.Font = (hoja1.Cell(2, 2).Style.Font);
                    }
                }

                #endregion

                #region Establece las formula de la sumatoria de los montos por mes

                celdaCabecera = 11;
                int renglonSumatoria = renglones + 1;
                for (int i = 11; i < (max + 11); i++)
                {

                    hoja1.Cell(renglonSumatoria, i).FormulaA1 = $"=SUM({ColumnIndexToColumnLetter(i)}2:{ColumnIndexToColumnLetter(i)}{renglones})";
                                       
                    hoja1.Cell(renglonSumatoria, i).Style.NumberFormat.Format = formatoMoneda;
                    hoja1.Cell(renglonSumatoria, i).Style.Font = (hoja1.Cell(1, 1).Style.Font);
                    hoja1.Cell(renglonSumatoria, i).Style.Fill.BackgroundColor = (workbook.Worksheet(1).Cell(1, 1).Style.Fill.BackgroundColor);
                  
                }
                #endregion

               

                hoja1.Columns().AdjustToContents();
                var ws = workbook.Worksheet(1);


                workbook.CalculateMode = XLCalculateMode.Auto;
                workbook.SaveAs(this.rutaArchivoSalida);

            }
            catch (Exception)
            {

                throw;
            }

        }


        static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        public DataSet leerExcel(MemoryStream msExcel)
        {
            DataSet result;
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(msExcel))
            {
                result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = this.useHeaderRow
                    }
                });

                if (!result.Tables.Contains(this.hojaInforme))
                {
                    throw new Exception($"La plantilla no contiene la Hoja '{this.hojaInforme}'");
                }

            }
            return result;
        }

        public string obtenerMes(int mes)
        {
            string textoMes = string.Empty;
            switch (mes)
            {
                case 1:
                    textoMes = "Enero";
                    break;
                case 2:
                    textoMes = "Febrero";
                    break;
                case 3:
                    textoMes = "Marzo";
                    break;
                case 4:
                    textoMes = "Abril";
                    break;
                case 5:
                    textoMes = "Mayo";
                    break;
                case 6:
                    textoMes = "Junio";
                    break;
                case 7:
                    textoMes = "Julio";
                    break;
                case 8:
                    textoMes = "Agosto";
                    break;
                case 9:
                    textoMes = "Septiembre";
                    break;
                case 10:
                    textoMes = "Octubre";
                    break;
                case 11:
                    textoMes = "Noviembre";
                    break;
                case 12:
                    textoMes = "Diciembre";
                    break;
            }

            return textoMes;
        }
    }
}
