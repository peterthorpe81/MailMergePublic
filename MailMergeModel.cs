using System.ComponentModel.DataAnnotations;
using Microsoft.Graph;

namespace MailMerge;

public class MailMergeModel
{
    public string? From { get; set; }
    public WorkbookTableColumn? ToField { get; set; }
    public WorkbookTableColumn? CcField { get; set; }
    public WorkbookTableColumn? BccField { get; set; }

    [MaxLength(100000)]
    public string? EmailBody { get; set;}
    [MaxLength(255)]
    public string? EmailSubject { get; set;}

    public WorkbookTable? Table { get; set; }
}
