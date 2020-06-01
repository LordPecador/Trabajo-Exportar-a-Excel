using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PROG2_ExportarExcel_PatricioAlmonacid_CarlosKlee_ErwinPerez
{
    class Alumprog2
    {
        private string rut;
        private string nombre;
        private int edad;
        private string seccion;
        private string asignatura;
        private int nota;

        public Alumprog2()
        {
        rut = "";
        nombre = "";
        edad = 0;
        seccion = "";
        asignatura = "";
        nota = 0;
        }

        public string Rut { get; set; }
        public string Nombre { get; set; }
        public int Edad { get; set; }
        public string Seccion { get; set; }
        public string Asignatura { get; set; }
        public int Nota { get; set; }

    }
}
