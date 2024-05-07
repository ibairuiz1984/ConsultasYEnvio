using OfficeOpenXml;


namespace ConsultasYEnvio
{
    class GeneradorExcel
    {
        public static void GenerarExcel(Persona persona)
        {
            // Crear un nuevo archivo Excel
            var fileInfo = new FileInfo(@"C:\Users\ibai.ruiz.ext\source\repos\ConsultasYEnvio\Persona.xlsx");
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Agregar una hoja al archivo
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Persona");

                // Escribir los datos de la persona en las celdas
                worksheet.Cells["A1"].Value = "Id";
                worksheet.Cells["B1"].Value = "Nombre";
                worksheet.Cells["C1"].Value = "Apellidos";
                worksheet.Cells["D1"].Value = "Edad";
                worksheet.Cells["E1"].Value = "Ocupación";

                worksheet.Cells["A2"].Value = persona.Id;
                worksheet.Cells["B2"].Value = persona.Nombre;
                worksheet.Cells["C2"].Value = persona.Apellidos;
                worksheet.Cells["D2"].Value = persona.Edad;
                worksheet.Cells["E2"].Value = persona.Ocupacion;

                // Guardar el archivo Excel
                package.Save();
            }
        }
    }
}
