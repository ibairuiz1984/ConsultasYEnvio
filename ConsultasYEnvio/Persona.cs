using System;

namespace ConsultasYEnvio
{
    class Persona
    {
        public string Nombre { get; set; }
        public string Apellidos { get; set; }
        public int Edad { get; set; }
        public string Ocupacion { get; set; }
        public int Id { get; set; }

        public Persona(int id, string nombre, string apellidos, int edad, string ocupacion)
        {
            Id = id;
            Nombre = nombre;
            Apellidos = apellidos;
            Edad = edad;
            Ocupacion = ocupacion;
        }
    }
}
