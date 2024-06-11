using System;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;
using File = System.IO.File;

namespace AttachmentMapping
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://sonyapc.sharepoint.com/sites/S002-iHR-Archival/HR/";
            string listName = "TH_Leave_Data";
            string username = "Connectadmin@sony.onmicrosoft.com";
            string password = "THX@v0lum3";

            SecureString securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            string localFolderPath = @"C:\Users\7000036422\Pictures\TH\leave\FY24-1"; // Specify the local folder path here

            // Connect to SharePoint site
            using (ClientContext context = new ClientContext(siteUrl))
            {
                // Provide credentials
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);

                // Get the list
                List list = context.Web.Lists.GetByTitle(listName);

                // Define CAML query to filter items where detail_line equals 1
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View><RowLimit>100</RowLimit></View>"; // Set the initial batch size

                ListItemCollectionPosition position = null;

                do
                {
                    query.ListItemCollectionPosition = position;
                    ListItemCollection items = list.GetItems(query);

                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        try
                        {
                            // Get claim_id from the current item
                            string claimId = item["application_id"].ToString();

                            // Check if the folder exists
                            if (Directory.Exists(localFolderPath))
                            {
                                // Get files in the folder
                                string[] files = Directory.GetFiles(localFolderPath);

                                // Loop through each file in the folder
                                foreach (string filePath in files)
                                {
                                    // Check if the file name starts with the desired claim ID
                                    if (Path.GetFileName(filePath).StartsWith(claimId))
                                    {
                                        Console.WriteLine($"Found document '{Path.GetFileName(filePath)}' in the local folder.");

                                        // Read the file content from local drive
                                        byte[] fileContent = File.ReadAllBytes(filePath);

                                        // Add attachment to list item
                                        AttachmentCreationInformation attachmentInfo = new AttachmentCreationInformation
                                        {
                                            FileName = Path.GetFileName(filePath),
                                            ContentStream = new MemoryStream(fileContent)
                                        };

                                        Attachment attachment = item.AttachmentFiles.Add(attachmentInfo);
                                        context.ExecuteQuery();

                                        Console.WriteLine("Attachment added successfully.");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Local folder does not exist.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error adding attachment for item ID: {item.Id}");
                            Console.WriteLine($"Error message: {ex.Message}");
                        }
                    }
                     
                    position = items.ListItemCollectionPosition;
                } while (position != null);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }
    }
}
