using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Web;

public static class Email
{
    static MailAddress fromAddress = new MailAddress("script@server.com", "Automated Script");
    static SmtpClient smtpClient = new SmtpClient("server@server.com", 25);
    static MailAddress replayToAddress = new MailAddress("email@domain.com", "Sender Name");

    public static bool SendEmailToAddress(string toAddress, string emailSubject, string emailBody)
    {
        MailMessage mailMessage = new MailMessage();
        mailMessage.To.Add(toAddress);
        mailMessage.From = fromAddress;
        mailMessage.Subject = emailSubject;
        mailMessage.Body = emailBody;
        mailMessage.ReplyToList.Add(replayToAddress);
        mailMessage.IsBodyHtml = true;
        return SendEmailFinal(mailMessage);
    }

    public static bool SendEmailWithAttachment(string toAddress, string emailSubject, string emailBody, Attachment attachment)
    {
        MailMessage mailMessage = new MailMessage();
        mailMessage.To.Add(toAddress);
        mailMessage.From = fromAddress;
        mailMessage.Subject = emailSubject;
        mailMessage.Body = emailBody;
        mailMessage.ReplyToList.Add(replayToAddress);
        mailMessage.IsBodyHtml = true;
        mailMessage.Attachments.Add(attachment);
        return SendEmailFinal(mailMessage);
    }

    public static bool SendEmailWithAttachment(string toAddress, string emailSubject, string emailBody, string attachmentFilePath)
    {
        MailMessage mailMessage = new MailMessage();
        mailMessage.To.Add(toAddress);
        mailMessage.From = fromAddress;
        mailMessage.Subject = emailSubject;
        mailMessage.Body = emailBody;
        mailMessage.ReplyToList.Add(replayToAddress);
        mailMessage.IsBodyHtml = true;
        Attachment attachment = new Attachment(attachmentFilePath);
        mailMessage.Attachments.Add(attachment);
        return SendEmailFinal(mailMessage);
    }

    private static bool SendEmailFinal(MailMessage mailMessage)
    {
        try
        {
            smtpClient.Send(mailMessage);
            return true;
        }
        catch (Exception e)
        {
            // Email sending failed
            return false;
        }
    }
}
