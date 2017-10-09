using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;

namespace CS01
{
    class Email
    {
        public string HOST = "smtp.gmail.com";
        public string FROM = "lddsoftware501@gmail.com";
        public string TO = "lddsoftware501@gmail.com";
        public int PORT = 587;
        public string USER = "lddsoftware501@gmail.com";
        public string PASSWORD = "89@l24%D01d?501";


        public void sendEmail(string _body)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtp = new SmtpClient(HOST);

                mail.From = new MailAddress(FROM);
                mail.To.Add(TO);
                mail.Subject = "[BG] Annapolis Status";
                mail.Body = _body;

                smtp.Port = PORT;
                smtp.Credentials = new System.Net.NetworkCredential(USER, PASSWORD);
                smtp.EnableSsl = true;

                smtp.Send(mail);
            } 
            catch(Exception e) 
            {
                Console.WriteLine("Error: "+ e.Message);
            }
        }
    }
}
