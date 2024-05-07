using OfficeOpenXml;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace ConsultasYEnvio
{
    class Program
    {
        static void Main(string[] args)
        {
            // Establecer el contexto de licencia para EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Solicitar datos de la persona
            Console.WriteLine("Ingrese el nombre de la persona:");
            string nombre = Console.ReadLine();

            // Validar que el nombre no contenga números
            while (ContieneNumeros(nombre))
            {
                Console.WriteLine("El nombre no puede contener números. Por favor, ingrese un nombre válido:");
                nombre = Console.ReadLine();
            }
            // Validar que los apellidos no contengan números
            Console.WriteLine("Ingrese los apellidos de la persona:");
            string apellidos = Console.ReadLine();

            while (ContieneNumeros(apellidos))
            {
                Console.WriteLine("Ingrese los apellidos de la persona:");
                apellidos = Console.ReadLine();
            }

            Console.WriteLine("Ingrese la edad de la persona:");
            int edad;
            // Validar que se ingrese un número válido para la edad
            while (!int.TryParse(Console.ReadLine(), out edad) || edad <= 0)
            {
                Console.WriteLine("Por favor, ingrese una edad válida:");
            }

            Console.WriteLine("Ingrese la ocupación de la persona:");
            string ocupacion = Console.ReadLine();

            // Generar un ID aleatorio de 5 dígitos
            Random random = new Random();
            int id = random.Next(10000, 99999); // Generar un número aleatorio entre 10000 y 99999

            // Crear la instancia de la persona
            Persona persona = new Persona(id, nombre, apellidos, edad, ocupacion);

            // Eliminar el archivo Excel anterior si existe
            string archivoAnterior = @"C:\Users\ibai.ruiz.ext\source\repos\ConsultasYEnvio\Persona.xlsx";
            if (File.Exists(archivoAnterior))
            {
                File.Delete(archivoAnterior);
            }

            // Generar el archivo Excel
            GeneradorExcel.GenerarExcel(persona);
            Console.WriteLine("Archivo Excel generado con éxito.");

            // Ruta del archivo Excel generado
            string archivoExcel = @"C:\Users\ibai.ruiz.ext\source\repos\ConsultasYEnvio\Persona.xlsx";

            // Envío del correo electrónico con el archivo adjunto al correo predeterminado
            EnvioMail.EnviarCorreo("example@hotmail.com", "Archivo Excel generado", "Adjunto encontrarás el archivo Excel generado.", archivoExcel);

            Console.WriteLine("Correo electrónico enviado con éxito.");
        }
        // Función para verificar si una cadena contiene números
        static bool ContieneNumeros(string cadena)
        {
            foreach (char c in cadena)
            {
                if (char.IsDigit(c))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
