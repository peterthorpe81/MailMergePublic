using System.ComponentModel.DataAnnotations;
using Microsoft.Graph;
using System.Text.RegularExpressions;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Web;

namespace MailMerge;
public partial class MailMergedRecord
{
    public static async Task<MailMergedRecord> Create(MailMergeModel template, WorkbookTableRow row)
    {
        var merge = new MailMergedRecord(template, row);
        await merge.Initialize();
        return merge;
    }

    [GeneratedRegex(@"\(\(.*?\)\)", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 10000)]
    private static partial Regex MergeFields();

    private static EmailAddressAttribute emailValid = new EmailAddressAttribute();

    public List<Exception> Exceptions { get; private set; }
    private MailMergedRecord(MailMergeModel template, WorkbookTableRow row)
    {        
        Row = row;
        Template = template;
        Exceptions = new List<Exception>();
    }
    private async Task Initialize()
    {
        //possibly async in future
        await Task.Yield();

        object[][] rowvaluesx =    JsonSerializer.Deserialize<object[][]>(Row.Values)!;
        object[] rowvalues = rowvaluesx[0];

  
        if (Template.From is null || !emailValid.IsValid(Template.From))                        
            Exceptions.Add(new ArgumentException($"Invalid From: Email {Template.From} Row: {Row.Index + 1}"));
        else
            From = Template.From;

        if (Template.ToField is not null)
            To = AddAddresses(CellValue(Template.ToField, rowvalues), Row.Index, true);  
        else            
            Exceptions.Add(new ArgumentException($"Missing To: Email Address Row: {Row.Index + 1}"));

        if (Template.CcField is not null)
            Cc = AddAddresses(CellValue(Template.CcField, rowvalues), Row.Index);
            
        if (Template.BccField is not null)
            Bcc = AddAddresses(CellValue(Template.BccField, rowvalues), Row.Index);


        if (Template.EmailBody is not null)
            MergedBody = MergeFields().Replace(Template.EmailBody, delegate(Match match)
            {
                string key = match.Groups[0].Value;
                return CellValue(key.Substring(2,key.Length-4), rowvalues);
            });
        else            
            Exceptions.Add(new ArgumentException($"Email Body Is Empty"));

        
        if (Template.EmailSubject is not null)
            MergedSubject = MergeFields().Replace(Template.EmailSubject, delegate(Match match)
            {
                string key = match.Groups[0].Value;
                return CellValue(key.Substring(2,key.Length-4), rowvalues);
            });
        else            
            Exceptions.Add(new ArgumentException($"Subject Is Empty"));

        
    } 

    public string CellValue(string columnName, object[] row)
    {        
        ArgumentNullException.ThrowIfNull(Template?.Table?.Columns);
        var column = Template.Table.Columns.Where(x => x.Name == columnName).SingleOrDefault();
        if (column is null)
        {
            Exceptions.Add(new ArgumentException($"Column {columnName} Not Found"));
            return String.Empty;
        }
        return HttpUtility.HtmlEncode(row[Template.Table.Columns.IndexOf(column)]?.ToString() ?? String.Empty);
    }

    public string CellValue(WorkbookTableColumn column, object[] row)
    {
        ArgumentNullException.ThrowIfNull(Template?.Table?.Columns);
        return HttpUtility.HtmlEncode(row[Template.Table.Columns.IndexOf(column)]?.ToString() ?? String.Empty);
    }

    public MailMergeModel Template { get; private set;}
    public WorkbookTableRow Row { get; private set;}
    public string From { get; private set; } = null!;
    public List<Recipient> To { get; private set;} = null!;
    public List<Recipient> Cc { get; private set;} = null!;
    public List<Recipient> Bcc { get; private set;} = null!;
    public string MergedBody { get; set;} = null!;
    public string MergedSubject { get; set;} = null!;
    public bool Valid { get => Exceptions.Count == 0 ? true : false; }

    public string ToList
    {
        get 
        {
            if (To is null)
                return String.Empty;
            string tmp = "";
            foreach (var ad in To)tmp += ad.EmailAddress.Address + ";";
            return tmp;
        }
    } 

     public string CcList
    {
        get 
        {
            if (Cc is null)
                return String.Empty;
            string tmp = "";
            foreach (var ad in Cc)tmp += ad.EmailAddress.Address+ ";";
            return tmp;
        }
    } 

     public string BccList
    {
        get 
        {
            if (Bcc is null)
                return String.Empty;
            string tmp = "";
            foreach (var ad in Bcc)tmp += ad.EmailAddress.Address+ ";";
            return tmp;
        }
    } 
    public Microsoft.Graph.Message GetMessage()
    {
        return new Microsoft.Graph.Message
            {
                Subject = MergedSubject,
                Body = new ItemBody
                {
                    ContentType = Microsoft.Graph.BodyType.Html,
                    Content = MergedBody
                },
                ToRecipients = To,
                CcRecipients = Cc,
                BccRecipients = Bcc
            };
    }


    private List<Recipient> AddAddresses(string addresses, int? RowIndex, bool required = false)
    {
        List<Recipient> recipients = new List<Recipient>();
        foreach (var address in addresses.Split(';').Select(p => p.Trim()))
        {
            if (emailValid.IsValid(address))
            {
                recipients.Add(new Recipient
                {
                    EmailAddress = new Microsoft.Graph.EmailAddress
                    {
                        Address = address
                    }
                });
            }
            else
            {
                Exceptions.Add(new ArgumentException($"Invalid Email {address} Row: {Row.Index + 1}"));
            }
        }
        if (required && recipients.Count == 0)
            Exceptions.Add(new ArgumentException($"Missing Required Email Address Row: {Row.Index + 1}"));

        return recipients;
    }
}
