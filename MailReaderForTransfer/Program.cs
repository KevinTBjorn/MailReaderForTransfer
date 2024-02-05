using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using System.Configuration;
using System.Text.RegularExpressions;

namespace MailReaderForTransfer
{
    class MailData
    {
        public decimal DKK { get; set; }
        public decimal TZS { get; set; }
        public DateTime Date { get; set; }
        public string Company { get; set; }
        public string Recipient { get; set; }
        public int transaktionsnummer { get; set; }
        public string Description { get; set; }
    }

    internal class Program
    {
        private static string path = Directory.GetCurrentDirectory();
        public static List<MailData> transactionData = new List<MailData>();

        static void Main(string[] args)
        {
            // If using Professional version, put your serial key below.
            //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var email = ConfigurationManager.AppSettings["email"];
            var password = ConfigurationManager.AppSettings["password"];

            List<MailData> transactionData = new List<MailData>();
            try
            {
                // Create IMAP client.
                using (var client = new ImapClient())
                {
                    client.Connect("imap-mail.outlook.com");
                    client.Authenticate(email, password);

                    // Select INBOX folder.
                    var WorldRemit = client.GetFolder("WorldRemit");
                    WorldRemit.Open(FolderAccess.ReadOnly);
                    var query = SearchQuery.SubjectContains("rt! Din transaktion").And(SearchQuery.FromContains("no-reply@info.worldremit.com"));
                    var uids = WorldRemit.Search(query);
                    foreach (var uid in uids)
                    {
                        var message = WorldRemit.GetMessage(uid);

                        // write the message to a file
                        Console.WriteLine("{0} \n", string.Format(message.Subject.Substring(28, 9)));
                        Console.WriteLine("{0} \n", message.Date.DateTime);
                        string[] lines = message.TextBody.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                        Console.WriteLine(lines[1]); //firma
                        var match = Regex.Match(lines[6], @"(?i)dig:\s+(.+?)\s+har");
                        if (match.Success)
                        {
                            Console.WriteLine(match.Groups[1].Value);
                        }
                        Console.WriteLine(lines[12].Substring(30).Split(new char[] { ' ' })[0]); //modtageren
                        Console.WriteLine(Convert.ToDecimal(lines[12].Substring(19).Split(new char[] { ' ' })[0]) / 100); //DKK
                        Console.WriteLine(Convert.ToDecimal(lines[13]) / 100); //TZS
                        foreach (var line in lines)
                        {
                            Console.WriteLine(line);
                        }
                    }

                    client.Disconnect(true);
                }

                Console.WriteLine("\n færdig");
                Console.ReadLine();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}