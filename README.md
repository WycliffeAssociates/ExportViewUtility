# ExportViewUtility
This was an internal tool to fufill the need of emailing a csv file of all of the records in a view.
The tool uses a fetchxml query to gather the data and then sends them through a configured smtp server

#Building
Standard build process either use Visual Studio or MSBUILD to build the application

#Usage
This is a command line tool so you are going to need to run it from there. You also
needs both a configuration file and a xml file containing the fetchXml you want to execute.
After you have both of these things run the application via `ExportViewUtility.exe <jsonconfigfile>`

#Configuration

Here is an example configuration file
`
{
    "crmConnectionString":"crm connection string",
    "fetchXmlFile":"fetchfile.xml",
    "columnMapping":[
        {
            "field":"fieldOne",
            "label":"Output File Label For fieldOne"
        },
        {
            "field":"fieldTwo",
            "label":"Output File Label For fieldTwo"
        }
    ],
    "emailAddress":["recepient@server.com","secondrecepient@server.com"],
    "mailServer":{
        "host":"smtpserver",
        "port":587,
        "userName": "usernameForSMTPServer",
        "password": "passwordForSMTPServer",
        "emailAddress":"emailAddressToBeSentFrom@example.com"
    },
    "subject":"Subject of the email",
    "fileName":"NameOfFile.csv",
    "body":"Body of the email"
}
`
#Contributing
Contributions are welcome via pull requests (Improving this README is a real easy. Hint, Hint)
