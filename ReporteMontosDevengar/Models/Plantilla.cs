using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReporteMontosDevengar.Models
{
    class Plantilla
    {
        public string ExpedienteProyecto { get; set; }
        public string PropietarioEP { get; set; }
        public string Cuenta { get; set; }
        public string ReferenciaElara { get; set; }
        public string Producto {get; set; }
        public DateTime FechaActual { get; set; }
        public DateTime FechaFinIngreso { get; set; }
        public int MesesPorDevengar { get; set; }
        public Decimal MontoPorDevengar { get; set; }

    }
}
