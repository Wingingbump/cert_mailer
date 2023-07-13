using Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace cert_mailer;

public class EmailBuilder
{
    public string Recipient { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public string CertificatePath { get; set; }

    public EmailBuilder(string recipient, string courseName, string courseId, string certificatePath)
    {
        Recipient = recipient;
        Subject = SubjectBuilder(courseName, courseId);
        Body = BodyBuilder(courseName, courseId);
        CertificatePath = certificatePath;

    }

    public void CreateDraft()
    {
        Outlook.Application outlook = GetOutlookApplication();
        MailItem mail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
        mail.To = Recipient;
        mail.Subject = Subject;
        mail.HTMLBody = Body;
        mail.Attachments.Add(CertificatePath, OlAttachmentType.olByValue, Type.Missing, Type.Missing); // Will throw exception if file path is invalid
        mail.Save();
        mail.SentOnBehalfOfName = "certs@bmra.com";

        // Check if draft has been created
        if (mail.EntryID == null)
        {
            throw new System.Exception($"Failed to create draft for recipient: {Recipient}");
        }
    }

    private string BodyBuilder(string course, string BMRARef)
    {
        var defaultMessage = "Hello, <br><br>" +
             "Congratulations. Attached to this email is your “Certificate of Training” for your " +
             course +
             " course; BMRA reference number " +
             BMRARef +
             ". If you require any further assistance, please feel free to contact us." +
             "<br><br>Thank you," +
             "<br><br>" + SignatureBuilder();

        return defaultMessage;
    }

    private string SubjectBuilder(string course, string BMRARef)
    {
        var defaultmessage = "BMRA Ref: " +
            BMRARef +
            " /// Certificate of Training - " +
            course;
        return defaultmessage;
    }

    private string GetUsername()
    {
        Outlook.Application outlook = GetOutlookApplication();
        NameSpace outlookNamespace = outlook.GetNamespace("MAPI");
        var username = outlookNamespace.CurrentUser.Name;
        return username;
    }

    private string GetAddress()
    {
        Outlook.Application outlook = GetOutlookApplication();
        NameSpace outlookNamespace = outlook.GetNamespace("MAPI");
        AddressEntry currentUser = outlookNamespace.CurrentUser.AddressEntry;
        var smtpAddress = string.Empty;

        if (currentUser != null && currentUser.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
        {
            ExchangeUser exchUser = currentUser.GetExchangeUser();
            if (exchUser != null)
            {
                smtpAddress = exchUser.PrimarySmtpAddress;
            }
        }

        return smtpAddress;
    }

    private string SignatureBuilder()
    {
        var signatureTemplate = GetTemplateContent("cert_mailer.template.htm");
        signatureTemplate = signatureTemplate.Replace("{{Username}}", "<b>" + GetUsername() + "</b>");
        signatureTemplate = signatureTemplate.Replace("{{Email}}", GetAddress());
        return signatureTemplate;
    }


    private string GetTemplateContent(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceStream = assembly.GetManifestResourceStream(resourceName);

        if (resourceStream == null)
        {
            throw new ArgumentException($"Resource {resourceName} not found in assembly.");
        }

        using var reader = new StreamReader(resourceStream);
        return reader.ReadToEnd();
    }


    private Outlook.Application GetOutlookApplication()
    {
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");

        if (outlookType == null)
        {
            outlookType = Type.GetTypeFromProgID("Outlook.Application.16");
        }

        if (outlookType == null)
        {
            throw new System.Exception("Outlook is not installed on this computer.");
        }

        var outlook = Activator.CreateInstance(outlookType);
        Outlook.Application? outlookApp = null;

        if (outlook != null)
        {
            outlookApp = outlook as Outlook.Application;
        }

        if (outlookApp == null)
        {
            throw new System.Exception("Failed to create an instance of Outlook application.");
        }

        return outlookApp;
    }

}
