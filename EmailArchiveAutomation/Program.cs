using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SharpCompress.Archives;
using SharpCompress.Archives.Zip;
using SharpCompress.Common;

namespace EmailArchiveAutomation
{
    /// <summary>
    /// Automates the workflow: Fetch -> Filter -> Download -> Decrypt -> Rename -> Save to NAS.
    /// Requires NuGet: MailKit, SharpCompress, MimeKit
    /// </summary>
    class Program
    {
        // Configuration - Move to App.config/Secret Manager in production
        private const string ImapServer = "imap.gmail.com";
        private const int ImapPort = 993;
        private const string EmailUser = "email";
        private const string EmailPass = "password";

        private const string NasPath = @"\\NAS-SERVER\Archive\ProcessedDocs\";
        private const string TempFolder = @"C:\Temp\EmailAutomation\";

        static async Task Main(string[] args)
        {
            Directory.CreateDirectory(TempFolder);
            Console.WriteLine($"[{DateTime.Now}] Starting Automation Service...");

            try
            {
                using (var client = new ImapClient())
                {
                    // 1. Connect and Authenticate
                    await client.ConnectAsync(ImapServer, ImapPort, true);
                    await client.AuthenticateAsync(EmailUser, EmailPass);

                    // 2. Open Inbox and Filter Relevant Emails
                    var inbox = client.Inbox;
                    await inbox.OpenAsync(FolderAccess.ReadWrite);

                    // Search for unread emails from specific senders or with specific subjects
                    var query = SearchQuery.NotSeen.And(SearchQuery.SubjectContains("Secure Invoice"));
                    var uids = await inbox.SearchAsync(query);

                    Console.WriteLine($"Found {uids.Count} relevant emails to process.");

                    foreach (var uid in uids)
                    {
                        var message = await inbox.GetMessageAsync(uid);
                        await ProcessEmailAsync(message);

                        // Mark as seen/archived after processing
                        await inbox.AddFlagsAsync(uid, MessageFlags.Seen, true);
                    }

                    await client.DisconnectAsync(true);
                }
            }
            catch (Exception ex)
            {
                LogEvent("SYSTEM_CRITICAL", $"Global Failure: {ex.Message}");
            }
        }

        static async Task ProcessEmailAsync(MimeMessage message)
        {
            Console.WriteLine($"Processing: {message.Subject}");

            // 3. Detect Password from Body
            // Often passwords are sent in the body or follow a known logic (e.g., last 4 of ID)
            string detectedPassword = ExtractPasswordFromBody(message.TextBody);

            foreach (var attachment in message.Attachments)
            {
                if (attachment is MimePart part)
                {
                    string fileName = part.FileName;
                    string fullTempPath = Path.Combine(TempFolder, fileName);

                    // 4. Download Attachment
                    using (var stream = File.Create(fullTempPath))
                    {
                        await part.Content.DecodeToAsync(stream);
                    }

                    // 5. Decrypt File (Assuming ZIP for this example)
                    if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        DecryptAndSave(fullTempPath, detectedPassword, message);
                    }
                }
            }
        }

        static string ExtractPasswordFromBody(string body)
        {
            // Regex to find patterns like "Pass: 12345" or similar
            var match = Regex.Match(body, @"Password:\s*(\S+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : "default_password";
        }

        static void DecryptAndSave(string zipPath, string password, MimeMessage metadata)
        {
            try
            {
                using (var archive = ZipArchive.Open(zipPath, new ReaderOptions { Password = password }))
                {
                    foreach (var entry in archive.Entries.Where(e => !e.IsDirectory))
                    {
                        // 6. Rename File (Workflow logic: Date_Sender_OriginalName)
                        string cleanSender = Regex.Replace(metadata.From.ToString(), @"[^a-zA-Z0-9]", "_");
                        string newName = $"{DateTime.Now:yyyyMMdd}_{cleanSender}_{entry.Key}";
                        string finalPath = Path.Combine(NasPath, newName);

                        // 7. Save to NAS
                        entry.WriteToDirectory(TempFolder, new ExtractionOptions { ExtractFullPath = true, Overwrite = true });

                        string extractedFile = Path.Combine(TempFolder, entry.Key);
                        File.Move(extractedFile, finalPath, true);

                        // 8. Generate Log + Notification
                        LogEvent("SUCCESS", $"Archived {newName} to NAS.");
                    }
                }

                // Cleanup temp zip
                File.Delete(zipPath);
            }
            catch (Exception ex)
            {
                LogEvent("DECRYPTION_FAIL", $"Failed to decrypt {zipPath}: {ex.Message}");
            }
        }

        static void LogEvent(string status, string message)
        {
            string logEntry = $"[{DateTime.Now:G}] [{status}] {message}";
            Console.WriteLine(logEntry);
            File.AppendAllText(Path.Combine(NasPath, "automation_log.txt"), logEntry + Environment.NewLine);

            // Note: Integration with Slack/Teams Webhooks could be added here
        }
    }
}