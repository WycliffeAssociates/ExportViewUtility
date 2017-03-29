using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ExportViewUtility.Models;
using System.IO;
using System.Net.Mail;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System.Net;

namespace ExportViewUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            //Check to see if the configuration file exists
            if (args.Length == 0)
            {
                Console.WriteLine("Usage ExportViewUtility.exe <json config file>");
                return;
            }
            if (!File.Exists(args[0]))
            {
                Console.WriteLine($"Configuration File {args[0]} wasn't found");
                return;
            }

            //Load configuration File

            Configuration config = JsonConvert.DeserializeObject<Configuration>(File.ReadAllText(args[0]));

            //Connect to CRM
            CrmServiceClient service = new CrmServiceClient(config.crmConnectionString);


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
                    Console.WriteLine($"FetchXml File: {config.fetchXmlFile} wasn't found");
                    return;
                }

                query = File.ReadAllText(config.fetchXmlFile);

            }
            //Query CRM
            Console.WriteLine("Querying CRM");

            FetchExpression fetchQuery = new FetchExpression(query);
            EntityCollection result = service.RetrieveMultiple(fetchQuery);

            if (result.Entities.Count == 0)
            {
                Console.WriteLine($"The result didn't return any records");
            }

            //Generate List

            Console.WriteLine("Writing File");

            //Get the labels
            List<string> headerLabels = new List<string>();
            foreach (ColumnMapping i in config.columnMapping)
            {
                //Quote the labels so that they load into the csv correctly
                headerLabels.Add($@"""{i.label}""");
            }

            //Create file
            //I'm using Memory stream so that I don't actually have to write the file to disk
            MemoryStream tempFile = new MemoryStream();
            StreamWriter writer = new StreamWriter(tempFile);

            //Write header
            writer.WriteLine(string.Join(",",headerLabels));

            List<string> row;

            foreach (Entity e in result.Entities)
            {
                //Clear the row
                row = new List<string>();

                foreach (ColumnMapping c in config.columnMapping)
                {
                    //If the entity contains the value insert it. If not then insert a blank value
                    if (e.Contains(c.field))
                    {
                        object value;

                        //If it is an aliased field then cast that value
                        if (e[c.field] is AliasedValue)
                        {
                            value = ((AliasedValue)e[c.field]).Value;
                        }
                        else
                        {
                            value = e[c.field];
                        }

                        //Add the quoted row
                        row.Add($@"""{value.ToString()}""");
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
            SmtpClient client = new SmtpClient(config.mailServer.host,config.mailServer.port);
            client.Credentials = new NetworkCredential(config.mailServer.userName, config.mailServer.password);
            client.EnableSsl = true;

            //Create the mail message
            MailMessage message = new MailMessage();
            message.From = new MailAddress(config.mailServer.emailAddress);

            //Add all the to messages
            foreach (string i in config.emailAddress)
            {
                message.To.Add(new MailAddress(i));
            }

            message.Subject = config.subject;
            message.Body = config.body;

            message.Attachments.Add(new Attachment(tempFile, config.fileName));

            //Send the message
            Console.WriteLine("Sending file to " + string.Join(",", config.emailAddress));
            client.Send(message);

        }
    }
}
