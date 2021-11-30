using Microsoft.AspNet.SignalR.Client;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace POSTTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter folder path to write PST Mail Items to(i.e.c:\\temp\\Email): ");

            String OutputRootPath = Console.ReadLine();
            if (System.IO.Directory.Exists(OutputRootPath) == false)
                return;
            WritePSTFilesToFolder(OutputRootPath);
            //var sms = "Hi, please DON'T REPLY This is a Test by C# Via SMS Eagle using Hash #";
            //    var smsParams = new SmsParams();
            //    smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
            //    smsParams.to = "+61437057198";
            //    smsParams.message = sms;

            //    var smsMain = new SmsMain();
            //    smsMain.method = "sms.send_sms";
            //    smsMain.params1 = smsParams;

            //    var result = await SendSmsAsync(smsMain);
            //    Console.WriteLine(result);
            //SendBulkSms();
        }

        private static void ProducePST(String newPstFileLocation)
        {
            Application app = new Application();
            NameSpace ns = app.GetNamespace("MAPI");
            MAPIFolder RootFolder = ns.PickFolder();

            Stores stores = null;
            Folder rootFolder = null;
            string storePath = string.Empty;

            try
            {
                ns.AddStoreEx(newPstFileLocation, OlStoreType.olStoreUnicode);
                stores = ns.Session.Stores;
                for (int i = 1; i <= stores.Count; i++)
                {
                    Store currStore = stores[i];
                    if (currStore.FilePath == newPstFileLocation)
                    {    
                        rootFolder = (Folder)currStore.GetRootFolder();
                        //ns.RemoveStore(rootFolder);
                        
                    }
                    if (currStore != null) Marshal.ReleaseComObject(currStore);
                }
                //var mail = ns.Session.OpenSharedItem(, MailItem);
            }
            finally
            {
                if (ns != null) Marshal.ReleaseComObject(ns);
                if (stores != null) Marshal.ReleaseComObject(stores);
                if (rootFolder != null) Marshal.ReleaseComObject(rootFolder);
            }
        }

        private static void WritePSTFilesToFolder(String OutputPath)
        {
            Application app = new Application();
            NameSpace outlookNs = app.GetNamespace("MAPI");
            MAPIFolder RootFolder = outlookNs.PickFolder();
            if (RootFolder != null) //It may be NULL if you press the Cancel button
                // Traverse through all folders in the PST file
                foreach (MAPIFolder SubFolder in RootFolder.Folders)
                {
                    Iterate(SubFolder, OutputPath);
                }
        }

        private static void Iterate(MAPIFolder RootFolder, String OutputPath)
        {
            //First, write any email items that may exist at root folder level
            OutputPath = OutputPath + RemoveFileNameSpecialChars(RootFolder.Name) + @"\";
            WriteEmails(RootFolder, OutputPath);

            //Recurse Subfolders
            foreach (MAPIFolder SubFolder in RootFolder.Folders)
            {
                Iterate(SubFolder, OutputPath);
            }
        }

        private static void WriteEmails(MAPIFolder Folder, String OutputPath)
        {
            Items items = Folder.Items;
            foreach (object item in items)
            {
                if (item is MailItem)
                {
                    // Retrieve the Object into MailItem
                    MailItem mailItem = item as MailItem;
                    Console.WriteLine("Saving message {0} .... into {1}", mailItem.Subject, OutputPath);
                    // Save the message to disk in MSG format
                    if (System.IO.Directory.Exists(OutputPath) == false)
                    {
                        System.IO.Directory.CreateDirectory(OutputPath);
                    }
                    try
                    {
                        String Subject = mailItem.Subject;
                        if (Subject == null)
                            Subject = "NULL";
                        String FilePathName = OutputPath + RemoveFileNameSpecialChars
                        (Subject.Replace('\u0009'.ToString(), "")) + ".msg"; //removes tab chars
                        mailItem.SaveAs(FilePathName, OlSaveAsType.olMSG);
                        mailItem.Move(OutputPath);
                        System.IO.File.SetCreationTime(FilePathName, mailItem.ReceivedTime);
                        System.IO.File.SetLastWriteTime(FilePathName, mailItem.ReceivedTime);
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        public static String RemoveFileNameSpecialChars(string FileName)
        {
            Regex regex = new Regex(@"[\w\s-'.,]");

            String validName = FileName;
            //identify invalid chars and replace those within the FileName string
            for (int i = 0; i < FileName.Length; i++)
            {
                Boolean matched = regex.IsMatch(FileName[i].ToString());
                if (matched == false)
                {
                    validName = validName.Replace(FileName[i].ToString(), "");
                }
            }

            return validName;
        }

        static void SendBulkSms()
        {
            var sms = " Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industrys standards dummy text ever since the 1500s, when an unknown printer took a galley of types";
            for (int i = 0; i < 10; i++)
            {
                var smsParams = new SmsParams();
                smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
                smsParams.to = "+61400252637";
                smsParams.message = (i + 1) + sms;

                var smsMain = new SmsMain();
                smsMain.method = "sms.send_sms";
                smsMain.params1 = smsParams;

                var result = SendSmsAsync(smsMain);
                Console.WriteLine(result);
            }

            sms = " make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised 1960s with the release Letraset";
            for (int i = 0; i < 10; i++)
            {
                var smsParams = new SmsParams();
                smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
                smsParams.to = "+61415866123";
                smsParams.message = (i + 1) + sms;

                var smsMain = new SmsMain();
                smsMain.method = "sms.send_sms";
                smsMain.params1 = smsParams;

                var result = SendSmsAsync(smsMain);
                Console.WriteLine(result);
            }

            sms = " It is a long established fact that a reader will be distracted by the readable content of page when looking at its layout. point of using Lorem Ipsum is that it has a more-or-less normal distribution";
            for (int i = 0; i < 10; i++)
            {
                var smsParams = new SmsParams();
                smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
                smsParams.to = "+61424255433";
                smsParams.message = (i + 1) + sms;

                var smsMain = new SmsMain();
                smsMain.method = "sms.send_sms";
                smsMain.params1 = smsParams;

                var result = SendSmsAsync(smsMain);
                Console.WriteLine(result);
            }

            sms = " as opposed to using, making it look like readable English. Many desktop publishing packages and web page editors now use Lorem Ipsum as their default model text, and a search for lorem ipsum will ur";
            for (int i = 0; i < 10; i++)
            {
                var smsParams = new SmsParams();
                smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
                smsParams.to = "+61416660909";
                smsParams.message = (i + 1) + sms;

                var smsMain = new SmsMain();
                smsMain.method = "sms.send_sms";
                smsMain.params1 = smsParams;

                var result = SendSmsAsync(smsMain);
                Console.WriteLine(result);
            }

            sms = " Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, Latin professor";
            for (int i = 0; i < 10; i++)
            {
                var smsParams = new SmsParams();
                smsParams.access_token = "u80vMmCBHIF6jXqPNparXKzyvSKlOif2";
                smsParams.to = "+61450255433";
                smsParams.message = (i + 1) + sms;

                var smsMain = new SmsMain();
                smsMain.method = "sms.send_sms";
                smsMain.params1 = smsParams;

                var result = SendSmsAsync(smsMain);
                Console.WriteLine(result);
            }
        }

        static async Task<string> SendSmsAsync(SmsMain smsMain)
        {
            var json = JsonConvert.SerializeObject(smsMain);
            json = json.Replace("params1", "params");
            var data = new StringContent(json, Encoding.UTF8, "application/json");

            var url = "http://192.168.18.43/jsonrpc/sms";
            using var client = new HttpClient();

            var response = await client.PostAsync(url, data);

            return await response.Content.ReadAsStringAsync();
        }
    }
}
