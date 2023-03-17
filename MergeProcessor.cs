using System.Linq;
using System.ComponentModel.DataAnnotations;
using Microsoft.Graph;

namespace MailMerge
{    
    public class MergeProcessor
    {        
        static SemaphoreSlim throttler = new SemaphoreSlim(initialCount: 1);

        public GraphServiceClient Client { get; set;}
        public bool MergeValid { get => MergeExceptions.Count == 0 ? true : false; }
        public bool SendValid { get => SendExceptions.Count == 0 ? true : false; }
        public bool Sending { get; set; }
        public List<Exception> MergeExceptions { get; private set; }
        public List<Exception> SendExceptions { get; private set; }
        public MailMergeModel Template { get; }
        public List<MailMergedRecord> MergedRecords { get; } = new List<MailMergedRecord>();

        private MergeProcessor(GraphServiceClient client, MailMergeModel template)
        {
            Client = client;
            Template = template;
            MergeExceptions = new List<Exception>();

            if (String.IsNullOrWhiteSpace(Template.EmailSubject))          
                MergeExceptions.Add(new ArgumentException($"Email Subject Is Empty"));

            if (String.IsNullOrWhiteSpace(Template.EmailBody))          
                MergeExceptions.Add(new ArgumentException($"Email Body Is Empty"));       

            if (Template.ToField is null)          
                MergeExceptions.Add(new ArgumentException($"Email To: is Empty"));       

            SendExceptions = new List<Exception>();
        }

        public static async Task<MergeProcessor> Create(GraphServiceClient client, MailMergeModel template)
        {
            var merge = new MergeProcessor(client, template);
            await merge.Initialize();
            return merge;
        }

        private async Task Initialize()
        {            
            ArgumentNullException.ThrowIfNull(Template?.Table?.Rows);
            if (MergeExceptions.Count > 0)
                return;

            foreach (var row in Template.Table.Rows)
            {
                var merge = await MailMergedRecord.Create(Template, row);
                MergedRecords.Add(merge);
                MergeExceptions.AddRange(merge.Exceptions);
            }
        }

        List<System.Threading.Tasks.Task> tasks = new List<System.Threading.Tasks.Task>();
            
        public async Task Send(IProgress<string>? progress = null)
        {    
            Sending = true;
            try
            {
                SendExceptions = new List<Exception>();
                
                if (!MergeValid)
                {
                    progress?.Report($"Invalid Merge");
                    throw new Exception("Invalid Merge");
                }   

                int sent = 0;
                int failed = 0;
                int total = MergedRecords.Count();

                foreach(var record in MergedRecords)
                {
                    await throttler.WaitAsync();
                    tasks.Add(System.Threading.Tasks.Task.Run(async () =>
                    {
                        try
                        {                        
                            await Client.Me
                                .SendMail(record.GetMessage(), true)
                                .Request()
                                .PostAsync();
                            sent++;
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            SendExceptions.Add(ex);
                        }
                        finally
                        {
                            throttler.Release();
                            progress?.Report($"Sent {sent} of {total}. {failed} Failed");
                        }
                    }));      
                    
                    progress?.Report($"Complete {sent} of {total}. {failed} Failed");                 
                }
            }
            finally
            {
                Sending = false;
            }

            return;
        }
            
    }
}
