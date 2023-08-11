using System;
using System.Collections.Generic;
using ExportViewUtility.Models;
using System.IO;
using System.Net.Mail;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System.Net;
using System.Text.Json;

namespace ExportViewUtility
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            //Check to see if the configuration file exists
            if (args.Length == 0)
            {
                Console.WriteLine("Usage ExportViewUtility.exe <json config file>");
                return;
            }
            if (!File.Exists(args[0]))
            {
                Console.WriteLine("Configuration File {0} wasn't found", args[0]);
                return;
            }

            //Load configuration File

            var config = JsonSerializer.Deserialize<Configuration>(File.OpenRead(args[0]));

            //Connect to CRM
            var service = new ServiceClient(config.crmConnectionString);


            //Get the query

            //If it was supplied then use it if not load the file
            string query;
            if (config.fetchXml != null)
            {
                query = config.fetchXml;
            }
            else
            {
                //Make sure the file that is supplied actually exists
                if (!File.Exists(config.fetchXmlFile))
                {
                    Console.WriteLine("FetchXml File: {0} wasn't found", config.fetchXmlFile);
                    return;
                }

                query = File.ReadAllText(config.fetchXmlFile);

            }
            //Query CRM
            Console.WriteLine("Querying CRM");

            var fetchQuery = new FetchExpression(query);
            var result = service.RetrieveMultiple(fetchQuery);

            // If no results are returned and the config states that we should skip it then we don't need to do anything else
            if (result.Entities.Count == 0 && config.skipIfEmpty)
            {
                return;
            }

            //Generate List

            Console.WriteLine("Writing File");

            //Get the labels
            var headerLabels = new List<string>();
            foreach (var i in config.columnMapping)
            {
                //Quote the labels so that they load into the csv correctly
                headerLabels.Add($@"""{i.label}""");
            }

            //Create file
            //I'm using Memory stream so that I don't actually have to write the file to disk
            var tempFile = new MemoryStream();
            var writer = new StreamWriter(tempFile);

            //Write header
            writer.WriteLine(string.Join(",",headerLabels));

            foreach (var e in result.Entities)
            {
                //Clear the row
                var row = new List<string>();

                foreach (var c in config.columnMapping)
                {
                    //If the entity contains the value insert it. If not then insert a blank value
                    if (e.Contains(c.field))
                    {
                        object value;

                        //If it is an aliased field then cast that value
                        if (e[c.field] is AliasedValue aliasedValue)
                        {
                            value = aliasedValue.Value;
                        }
                        else
                        {
                            value = e[c.field];
                        }

                        if (value is EntityReference reference)
                        {
                            value = reference.Name;
                        }

                        //Add the quoted row
                        row.Add(value == null ? @"""""" : $@"""{value}""");
                    }
                    else
                    {
                        row.Add("");
                    }
                }
                //Write the CSV row
                writer.WriteLine(string.Join(",", row));
            }

            //Flush out all pending writes before seeking
            writer.Flush();

            //Go back to the beginning of the file in order to read
            tempFile.Seek(0, SeekOrigin.Begin);


            //Email list

            //Initialize smtp client
            var client = new SmtpClient(config.mailServer.host,config.mailServer.port);
            client.Credentials = new NetworkCredential(config.mailServer.userName, config.mailServer.password);
            client.EnableSsl = true;

            //Create the mail message
            var message = new MailMessage();
            message.From = new MailAddress(config.mailServer.emailAddress);

            //Add all the to messages
            foreach (var i in config.emailAddress)
            {
                message.To.Add(new MailAddress(i));
            }

            message.Subject = config.subject;
            message.Body = config.body;

            message.Attachments.Add(new Attachment(tempFile, config.fileName));

            //Send the message
            Console.WriteLine("Sending file to {0}", string.Join(",", config.emailAddress));
            client.Send(message);

        }
    }
}
