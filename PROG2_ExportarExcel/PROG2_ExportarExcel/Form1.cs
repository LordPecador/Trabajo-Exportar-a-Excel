using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetLight; //libreria agregada
using System.IO;        //libreria agregada
using OfficeOpenXml;       //libreria agregada
using LicenseContext = OfficeOpenXml.LicenseContext;    //libreria agregada

using Microsoft.Office.Interop.Excel; // Uso de libreria office excel

using ObjExcel = Microsoft.Office.Interop.Excel; // creacion de variable para objeto excel (Aplication,Worbook y worksheet)


namespace PROG2_ExportarExcel_PatricioAlmonacid_CarlosKlee_ErwinPerez
{
    public partial class Form1 : Form
    {
        private List<Alumprog2> alumnos = new List<Alumprog2>(); // crear lista para la clase

        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // creacion de ruta donde guardaremos el excel

        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Alumprog2 alumno = new Alumprog2(); //creación de objeto (alumn) a la clase Alumprog2         

            alumno.Rut = textBox1.Text;
            alumno.Nombre = textBox2.Text;
            alumno.Edad = Int32.Parse(textBox3.Text);
            alumno.Seccion = textBox4.Text;
            alumno.Asignatura = textBox5.Text;
            alumno.Nota = Int32.Parse(textBox6.Text);

            alumnos.Add(alumno); //agrega la lista al objeto

            MessageBox.Show("REGISTRO GUARDADO");

            //limpia los campos de registro
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.dataGridView1.DataSource = null;
            this.dataGridView1.DataSource = alumnos; //agrega el contenido de la lista al dataGrid

            this.dataGridView1.Refresh();

            MessageBox.Show("MOSTRANDO LISTA");
         
        }

        private void Button3_Click(object sender, EventArgs e)
        {

        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
            textBox1.Text = alumnos[e.RowIndex].Rut;
            textBox2.Text = alumnos[e.RowIndex].Nombre;
            textBox3.Text = Convert.ToString(alumnos[e.RowIndex].Edad);
            textBox4.Text = alumnos[e.RowIndex].Seccion;
            textBox5.Text = alumnos[e.RowIndex].Asignatura;
            textBox6.Text = Convert.ToString(alumnos[e.RowIndex].Nota);

            //MessageBox.Show("MOSTRANDO REGISTRO");
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {

            // utilizamos la libreria SpreadsheetLight (exporta la data incluso si no tenemos un exel instalado)
            //agregamos una libreria para utilizar sus metodos
            SLDocument exporta = new SLDocument(); //creamos un objeto que permita usar la libreria agregada

            int iC = 1; //variable utilizada para contador
            foreach (DataGridViewColumn column in dataGridView1.Columns) //recorremos las columnas 
            {
                exporta.SetCellValue(1, iC, column.HeaderText.ToString());
                iC++;
            }

            int iR = 2;            
            foreach (DataGridViewRow row in dataGridView1.Rows) //recorremos el dataGrid por filas
            {
                exporta.SetCellValue(iR, 1, row.Cells[0].Value.ToString()); //row nos referimos a la fila
                exporta.SetCellValue(iR, 2, row.Cells[1].Value.ToString()); //aqui recorremos la fila 2 (exel comienza de 1 la pisicion)
                exporta.SetCellValue(iR, 3, row.Cells[2].Value.ToString());
                exporta.SetCellValue(iR, 4, row.Cells[3].Value.ToString());
                exporta.SetCellValue(iR, 5, row.Cells[4].Value.ToString());
                exporta.SetCellValue(iR, 6, row.Cells[5].Value.ToString());                
                iR++;               
            }
            exporta.SaveAs(@"C:\XLS\DataGrid.xlsx");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //condición para usar la libreria EPPlus

            var exel = new ExcelPackage(); //creamos objeto que permite usar la librería
            var pag = exel.Workbook.Worksheets.Add("Hoja data"); //creamos la pagina donde van los datos

            pag.Cells[1, 1].LoadFromCollection(alumnos, true); //a la pag carga la lista alumnos, con el encabezado, empezando de la posición 1f,1c

            FileInfo fileInfo = new FileInfo(@"C:\\XLS\Lista.xlsx"); //creamos la dirección donde guarda el archivo
            exel.SaveAs(fileInfo); //indicamos que guarde el archivo

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application objetoExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = objetoExcel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)objetoExcel.ActiveSheet;

            objetoExcel.Visible = true;

            ws.Cells[1, 1] = "Rut";
            ws.Cells[1, 2] = "Nombre";
            ws.Cells[1, 3] = "Edad";
            ws.Cells[1, 4] = "Seccion";
            ws.Cells[1, 5] = "Asignatura";
            ws.Cells[1, 6] = "Nota";

            ws.Cells[2, 1] = textBox1.Text;
            ws.Cells[2, 2] = textBox2.Text;
            ws.Cells[2, 3] = textBox3.Text;
            ws.Cells[2, 4] = textBox4.Text;
            ws.Cells[2, 5] = textBox5.Text;
            ws.Cells[2, 6] = textBox6.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ObjExcel.Application objAplication = new ObjExcel.Application();
            Workbook objLibro = objAplication.Workbooks.Add(XlSheetType.xlWorksheet); // Aca vamos agregar en nuestro libro de trabajo otro nuevo libro que sea hoja de calculos
            Worksheet objHoja = (Worksheet)objAplication.ActiveSheet; // Va hacer nuestra hoja de calculo dentro del aplicativo excel

            objAplication.Visible = true; // Este codigo es para que me Guarde el archivo excel  y si quiero que se abra automaticamente debo soleccionar truy

            //creacion de celdad de nuestro libro de trabajos mediante una extructura de arreglo la cual se indica las cordenadas 

            objHoja.Cells[1, 1] = "Nombre";
            objHoja.Cells[1, 2] = "Apellido";
            objHoja.Cells[1, 3] = "Telefono";
            objHoja.Cells[1, 4] = "Correo";

            objHoja.Cells[2, 1] = "Erwin";
            objHoja.Cells[2, 2] = "Perez";
            objHoja.Cells[2, 3] = "987633649";
            objHoja.Cells[2, 4] = "Erwin@gmail.com";

            objHoja.Cells[3, 1] = "Patricio";
            objHoja.Cells[3, 2] = "almonacid";
            objHoja.Cells[3, 3] = "987224640";
            objHoja.Cells[3, 4] = "PatricioAlmo@gmail.com";

            objHoja.Cells[4, 1] = "Carlos";
            objHoja.Cells[4, 2] = "Klee";
            objHoja.Cells[4, 3] = "941195490";
            objHoja.Cells[4, 4] = "Calafaro@gmail.com";

            objLibro.SaveAs("C:\\XLS\\ExcelData.xls"); // Realizamos el guardado de nuestro libro con el metodo SaveAs
            objLibro.Close(); //Se realiza el cierre del libro con el metodo close
            objAplication.Quit(); // se cierra la aplicacion con el metodo quit
        }
    }
}
