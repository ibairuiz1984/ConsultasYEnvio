using System.Net;
using System.Net.Mail;

namespace ConsultasYEnvio
{
    public static class EnvioMail
    {
        public static void EnviarCorreo(string destinatario, string asunto, string cuerpo, string archivoAdjunto)
        {
            // Creación del mensaje de correo electrónico
            MailMessage mensaje = new MailMessage("example@hotmail.com", destinatario, asunto, cuerpo);

            // Adjuntar el archivo Excel al mensaje
            Attachment adjunto = new Attachment(archivoAdjunto);
            mensaje.Attachments.Add(adjunto);

            // Configuración del cliente SMTP para Hotmail (Outlook.com)
            SmtpClient clienteSmtp = new SmtpClient("smtp-mail.outlook.com")
            {
                Port = 587,
                Credentials = new NetworkCredential("example@hotmail.com", "password"),
                EnableSsl = true
            };

            // Envío del mensaje de correo electrónico
            clienteSmtp.Send(mensaje);
        }
    }
}
