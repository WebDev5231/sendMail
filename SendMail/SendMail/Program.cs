using System;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using Dapper;

namespace SendMail
{
    class Program
    {
        static void Main(string[] args)
        {
            EnviarEmails();
        }

        static void EnviarEmails()
        {
            try
            {
                using (var connection = new SqlConnection("Data Source=(localdb)\\MSSQLLocalDB;Integrated Security=True;"))
                {
                    connection.Open();

                    connection.Execute("USE sendMail");

                    string query = "SELECT Email, Nome FROM Juiz";
                    var usuarios = connection.Query(query);

                    foreach (var usuario in usuarios)
                    {
                        Console.WriteLine($"Consultando e-mail do juiz {usuario.Nome}...");
                        EnviarEmail(usuario.Email, usuario.Nome);
                        Console.WriteLine($"E-mail para {usuario.Nome} enviado com sucesso! Aguardando para o próximo envio...");
                        Console.WriteLine("-------------------------------------------------------------");
                        System.Threading.Thread.Sleep(30000);
                    }
                }

                Console.WriteLine("Todos os e-mails foram enviados com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mails: {ex.Message}");
            }
        }

        static void EnviarEmail(string destinatario, string nome)
        {
            try
            {
                // Configurar o cliente SMTP para o Gmail
                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com");
                smtpClient.Port = 587; // Porta do servidor SMTP do Gmail
                smtpClient.EnableSsl = true;

                // Configurar o e-mail
                MailMessage mensagem = new MailMessage();
                mensagem.From = new MailAddress("viniciusgarciagomess@gmail.com"); // Substitua pelo seu endereço de e-mail do Gmail
                mensagem.To.Add(destinatario);

                // Montar o assunto do e-mail
                mensagem.Subject = $"Ex.mo(a) Juiz(a) Dr(a). {nome} – Assunto: Perito em Grafotécnica";

                mensagem.IsBodyHtml = true;
                mensagem.Body = $"<html><body style=\"line-height: 1.5\">" +
                                $"<p>Ex.mo(a) Juiz(a) Dr(a). {nome}, boa tarde.</p>" +
                                $"<p>Agradeço por sua atenção. Sou <b>Paulo Eduardo Porlan de Almeida</b>, Engenheiro de Produção aposentado, formado pela Escola Politécnica da Universidade de São Paulo. O objetivo deste contato é oferecer a V. Exa. a possibilidade de servi-lo(a), atuando como:</p>" +
                                $"<ul>" +
                                $"<li><b>Perito Grafotécnico</b></li>" +
                                $"<li><b>Perito em Documentoscopia</b></li>" +
                                $"<li><b>Perito em Papiloscopia</b></li>" +
                                $"</ul>" +
                                $"<p>Tendo realizado o meu cadastramento no Tribunal de Justiça deste estado, estou habilitado a colaborar em processos nos quais um perito numa dessas especialidades se mostre necessário. Ficarei honrado se puder servir a V. Exa., quando uma oportunidade se apresentar.</p>" +
                                $"<p>Em anexo, seguem meus documentos pessoais e certificados comprobatórios. Tenha um ótimo final de semana.</p>" +
                                $"<p><b>Cordialmente,</b></p>" +
                                $"<p><b>Paulo Eduardo Porlan de Almeida</b></p>" +
                                $"<p><b>Perito em Grafotécnica, Documentoscopia e Papiloscopia</b></p>" +
                                $"<p><b>Whatsapp: (64) 99211-7855</b></p>" +
                                $"</body></html>";

                // Anexar arquivos
                string[] filePaths = Directory.GetFiles(@"C:\fileMail");
                foreach (string filePath in filePaths)
                {
                    mensagem.Attachments.Add(new Attachment(filePath));
                }

                // Configurar as credenciais usando a senha de aplicativo
                smtpClient.Credentials = new NetworkCredential("viniciusgarciagomess@gmail.com", "ccmg qxzg erft gsvm"); // Substitua pela senha de aplicativo

                // Enviar o e-mail
                smtpClient.Send(mensagem);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao enviar e-mail para {nome}: {ex.Message}");
            }
        }
    }
}
