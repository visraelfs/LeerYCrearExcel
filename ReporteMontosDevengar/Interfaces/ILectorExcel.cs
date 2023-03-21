using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReporteMontosDevengar.Interfaces
{
    public interface ILectorExcel
    {
        DataSet leerExcel(MemoryStream msExcel);
    }
}
