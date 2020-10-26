using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Kernel.Geom;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.IO;
using iText.IO.Image;

namespace Pagina_Inicial
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void crearPDF()
        {


            PdfWriter pdfWrite = new PdfWriter("Reporte.pdf");
            PdfDocument pdf = new PdfDocument(pdfWrite);
            PageSize tamanioH = new PageSize(792, 612);
            Document documento = new Document(pdf, tamanioH);

            documento.SetMargins(60, 30, 55, 20);
            
            /*
            var parrafo = new Paragraph("Hola mundo");
            documento.Add(parrafo);
            documento.Close(); */
            PdfFont fontColumnas = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fontContenido = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            string[] columnas = { "Primer Tetramestre" };
            string asignaturas = "Asignaturas";
            
            float[] tamanios = { 5 };

            Table tabla1 = new Table(UnitValue.CreatePercentArray(tamanios));
            Table tabla = new Table(UnitValue.CreatePercentArray(tamanios));

            tabla1.SetWidth(UnitValue.CreatePercentValue(30));
            tabla.SetWidth(UnitValue.CreatePercentValue(30));


            tabla1.AddHeaderCell(new Cell().Add(new Paragraph(asignaturas).SetFont(fontColumnas)));

            foreach (string columna in columnas)
            {
                
                tabla.AddHeaderCell(new Cell().Add(new Paragraph(columna).SetFont(fontColumnas)));
                
            }

            string sql = "SELECT Nombre_M FROM materia LEft JOIN materia_carrera On materia.Cod_M = materia_carrera.Cod_m4 WHERE tetra = 1";

            //Conexion a la base de datos
            SqlConnection con = new SqlConnection(@"Data source = DESKTOP-JJ17LHE;Initial Catalog=Universidad;Integrated Security=true");
            con.Open();
            //Termina la conexion a la base de datos


            SqlCommand comando = new SqlCommand(sql, con);
            SqlDataReader reader = comando.ExecuteReader();

            while (reader.Read())
            {
                tabla.AddCell(new Cell().Add(new Paragraph(reader["Nombre_M"].ToString()).SetFont(fontContenido)));
            }
            documento.Add(new Paragraph("KARDEX"));

            documento.Add(tabla1);
            documento.Add(tabla);
            documento.Close();

            var logo = new iText.Layout.Element.Image(ImageDataFactory.Create("C:/Users/PORTEGE Z30-B/Desktop/logo-umm.jpg")).SetWidth(100).SetHeight(100);
            var plogo = new Paragraph("").Add(logo);

            var Titulo = new Paragraph("Universidad Metrorrey");
            Titulo.SetTextAlignment(TextAlignment.CENTER);
            Titulo.SetFontSize(25);

            var Carrera = new Paragraph("Arquitecto");
            Carrera.SetTextAlignment(TextAlignment.CENTER);
            Carrera.SetFontSize(20);

            var Alumno = new Paragraph("Alumno: ");
            Alumno.SetTextAlignment(TextAlignment.CENTER);
            Alumno.SetFontSize(20);

            var nombre = new Paragraph("Jose Miguel Rodriguez");
            nombre.SetUnderline(0.1f, -2f);
            nombre.SetTextAlignment(TextAlignment.CENTER);
            nombre.SetFontSize(20);

            var Matricula = new Paragraph("Matricula: ");
            Matricula.SetTextAlignment(TextAlignment.CENTER);
            Matricula.SetFontSize(20);

            var matricula = new Paragraph("207127");
            matricula.SetUnderline(0.1f, -2f);
            matricula.SetTextAlignment(TextAlignment.CENTER);
            matricula.SetFontSize(20);

            PdfDocument pdfDoc = new PdfDocument(new PdfReader("Reporte.pdf"), new PdfWriter("ReporteProducto.pdf"));

            Document doc = new Document(pdfDoc);

            int numeros = pdfDoc.GetNumberOfPages();

            for (int i = 1; i<=numeros; i++)
            {
                PdfPage pagina = pdfDoc.GetPage(i);

                float y = (pdfDoc.GetPage(i).GetPageSize().GetTop() - 15);
                doc.ShowTextAligned(plogo, 100, y - 15, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(Titulo, 600, y - 15, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(Carrera, 580, y - 50, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(Alumno, 470, y - 80, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(nombre, 620, y - 80, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(Matricula, 465, y - 110, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(matricula, 550, y - 110, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
                doc.ShowTextAligned(new Paragraph(string.Format("Pagina {0} de {1}", i, numeros)), pdfDoc.GetPage(i).GetPageSize().GetWidth() / 2, pdfDoc.GetPage(i).GetPageSize().GetBottom() + 30, i, TextAlignment.CENTER, VerticalAlignment.TOP, 0);
               
            }
            doc.Close();
            con.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            crearPDF();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
