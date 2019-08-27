using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MsgKit;
using Org.BouncyCastle.Bcpg.OpenPgp;
using PSTParse;
using PSTParse.MessageLayer;
using static System.Console;
using Newtonsoft.Json;
using PSTParse.ListsTablesPropertiesLayer;
using Attachment = PSTParse.MessageLayer.Attachment;
using Message = PSTParse.MessageLayer.Message;
using Recipient = PSTParse.MessageLayer.Recipient;

namespace PSTParseApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = new Stopwatch();
            sw.Start();
            var pstPath = @"/Users/jefe/projects/legalverse/libpff/examples/workingData/batch1-Tango.pst";
            var fileInfo = new FileInfo(pstPath);
            var pstSizeGigabytes = ((double)fileInfo.Length / 1000 / 1000 / 1000).ToString("0.000");
            using (var file = new PSTFile(pstPath))
            {
                Debug.Assert((double)fileInfo.Length / 1000 / 1000 / 1000 == file.SizeMB / 1000);
                //Console.WriteLine("Magic value: " + file.Header.DWMagic);
                //Console.WriteLine("Is Ansi? " + file.Header.IsANSI);

                var stack = new Stack<MailFolder>();
                stack.Push(file.TopOfPST);
                var totalCount = 0;
                //var maxSearchSize = 1500;
                var maxSearchSize = int.MaxValue;
                var totalEncryptedCount = 0;
                var totalUnsentCount = 0;
                var skippedFolders = new List<string>();
                var messagesCount = 0;
                var appointmentCount = 0;
                
                IList<object> pstItems = new List<object>();
                
                while (stack.Count > 0)
                {
                    var curFolder = stack.Pop();

                    foreach (var child in curFolder.SubFolders)
                    {
                        stack.Push(child);
                    }
                    var count = curFolder.Count;
                    var line = $"{string.Join(" -> ", curFolder.Path)}({curFolder.ContainerClass}) ({count} messages)";
                    if (curFolder.Path.Count > 1 && curFolder.ContainerClass != "" && curFolder.ContainerClass != "IPF.Note")
                    {
                        skippedFolders.Add(line);
                        continue;
                    }
                    WriteLine(line);

                    var currentFolderCount = 0;
                    string path = @"/Users/jefe/Desktop/attachments/";
                    Directory.CreateDirectory(path);

                    foreach (var ipmItem in curFolder.GetIpmItems())
                    {
                        string className = ipmItem.GetType().Name;
                        if(className != "Message" && className != "Meeting")
                            Console.WriteLine("{0} - {1}", className, ipmItem.MessageClass);

                        if (ipmItem is Message message)
                        {
                            message.GUID = Guid.NewGuid().ToString();
                            foreach (var attachment in message.Attachments)
                            {
                                attachment.SaveToDisk(Path.Combine(path, message.GUID));
                            }
                        }
                        
                        if (ipmItem is Meeting meeting)
                        {
                            meeting.GUID = Guid.NewGuid().ToString();
                            foreach (var attachment in meeting.Attachments)
                            {
                                attachment.SaveToDisk(Path.Combine(path, meeting.GUID));
                            }
                        } 
                            
                        
                        totalCount++;
                        currentFolderCount++;
                        pstItems.Add(ipmItem);
                        //if (totalCount >= 500)
                          //  break;
                    }
                    
                }
                sw.Stop();
                var elapsedSeconds = (double)sw.ElapsedMilliseconds / 1000;
                WriteLine("{0} messages total", totalCount);
                //WriteLine("{0} encrypted messages total", totalEncryptedCount);
                //WriteLine("{0} totalUnsentCount", totalUnsentCount);
                WriteLine("Parsed {0} ({1} GB) in {2:0.00} seconds", Path.GetFileName(pstPath), pstSizeGigabytes, elapsedSeconds);

                WriteLine("\r\nSkipped Folders:\r\n");
                foreach (var line in skippedFolders)
                {
                    WriteLine(line);
                }
                
                // Custom Json Options
                var serializerSettings = new JsonSerializerSettings();
                var jsonResolver = new PropertyRenameAndIgnoreSerializerContractResolver();
                
                // Messages
                jsonResolver.IgnoreProperty((typeof(Message)), "Data");
                jsonResolver.IgnoreProperty((typeof(Message)), "AttachmentPC");
                jsonResolver.IgnoreProperty((typeof(Message)), "RecipientTable");
                jsonResolver.IgnoreProperty((typeof(Meeting)), "IsRMSEncrypted");

                // Appointments
                jsonResolver.IgnoreProperty((typeof(Meeting)), "Data");
                jsonResolver.IgnoreProperty((typeof(Meeting)), "AttachmentPC");
                jsonResolver.IgnoreProperty((typeof(Meeting)), "RecipientTable");
                jsonResolver.IgnoreProperty((typeof(Meeting)), "IsRMSEncrypted");
                
                // Recipients
                jsonResolver.IgnoreProperty((typeof(Recipient)), "Tag");
                jsonResolver.IgnoreProperty((typeof(Recipient)), "Type");
                jsonResolver.IgnoreProperty((typeof(Recipient)), "ObjType");
                
                //Attachment
                jsonResolver.IgnoreProperty((typeof(Attachment)), "Data");
                
                // Contacts
                
                // Tasks
                
                jsonResolver.IgnoreProperty(typeof(BTHIndexNode), "Data");
                //jsonResolver.IgnoreProperty(typeof(PSTParse.LTP.TableContext), "RowMatrix");
                serializerSettings.ContractResolver = jsonResolver;
                
                File.WriteAllText(@"/Users/jefe/Desktop/ipmexport.json",JsonConvert.SerializeObject(pstItems, Formatting.Indented, serializerSettings));
                Read();
            }
        }
    }
}
