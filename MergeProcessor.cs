using System.Linq;
using System.ComponentModel.DataAnnotations;
using Microsoft.Graph;
using MailMerge.Shared;

namespace MailMerge
{    
    public class MergeProcessor
    {        
        static SemaphoreSlim throttler = new SemaphoreSlim(initialCount: 1);

        public GraphServiceClient Client { get; set;}
        public bool MergeValid { get => MergeExceptions.InnerExceptions.Count == 0 ? true : false; }
        public AggregateException MergeExceptions { get; private set; }
        public AggregateException SendExceptions { get; private set; }
        public MailMergeModel Template { get; }
        public List<MailMergedRecord> MergedRecords { get; } = new List<MailMergedRecord>();
        private MergeProcessor(GraphServiceClient client, MailMergeModel template)
        {
            Client = client;
            Template = template;
            MergeExceptions = new AggregateException();
            SendExceptions = new AggregateException();
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

            foreach (var row in Template.Table.Rows)
            {
                //IDictionary<string, string> expando = record;
                var merge = await MailMergedRecord.Create(Template, row);
                MergedRecords.Add(merge);
                MergeExceptions.InnerExceptions.Append(merge.Exceptions);
            }
        }
        List<System.Threading.Tasks.Task> tasks = new List<System.Threading.Tasks.Task>();
            
        public async Task Send(IProgress<string>? progress = null)
        {    
            SendExceptions = new AggregateException();
            
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
                        SendExceptions.InnerExceptions.Append(ex);
                    }
                    finally
                    {
                        throttler.Release();
                        progress?.Report($"Sent {sent} of {total}. {failed} Failed");
                    }
                }));      
                
                progress?.Report($"Complete {sent} of {total}. {failed} Failed");                 
            }

            return;
        }
            
    }
}
