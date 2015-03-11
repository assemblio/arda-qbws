using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services.Protocols;
using System.Xml.Serialization;
using System.Web.Services;
using System.Security.Cryptography;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Collections;
using System.Xml;
using MongoDB.Bson;
using MongoDB.Driver;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ArdaQBWS.Properties;
using System.Xml.Linq;
using System.Text;

namespace ArdaQBWS
{
    /// <summary>
    /// Summary description for QBWebService
    /// </summary>
    [WebService(
         Namespace = "http://developer.assemblio.com/",
         Name = "WCWebService",
         Description = "Quickbooks WebService in ASP.NET " +
                "QuickBooks WebConnector")]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class QBWebService : System.Web.Services.WebService
    {
        #region GlobalVariables
        System.Diagnostics.EventLog evLog = new System.Diagnostics.EventLog();
        public int count = 0;
        public ArrayList req = new ArrayList();
        #endregion

        [WebMethod]
        public string[] authenticate(string strUserName, string strPassword)
        {
            string[] authReturn = new string[2];

            // Generate a random session ticket 
            authReturn[0] = System.Guid.NewGuid().ToString();

          
            if (strUserName.Trim().Equals("test") && strPassword.Trim().Equals(Settings.Default.wsauthpassword))
            {
                // An empty string for authReturn[1] means asking QBWebConnector 
                // to connect to the company file that is currently openned in QB
                authReturn[1] = "";
            }
            else
            {
                authReturn[1] = "nvu";
            }

            return authReturn;
        }

        /// <summary>
        /// WebMethod# 4 - sendRequestXML()
        /// Signature: public string sendRequestXML(string ticket, string strHCPResponse, string strCompanyFileName, 
        /// string Country, int qbXMLMajorVers, int qbXMLMinorVers)
        /// 
        /// IN: 
        /// int qbXMLMajorVers
        /// int qbXMLMinorVers
        /// string ticket
        /// string strHCPResponse 
        /// string strCompanyFileName 
        /// string Country
        /// int qbXMLMajorVers
        /// int qbXMLMinorVers
        ///
        /// OUT:
        /// string request
        /// Possible values: 
        /// - “any_string” = Request XML for QBWebConnector to process
        /// - "" = No more request XML 
        /// </summary>
        [WebMethod(Description = "This web method facilitates web service to send request XML to QuickBooks via QBWebConnector", EnableSession = true)]
        public string sendRequestXML(string ticket, string strHCPResponse, string strCompanyFileName,
            string qbXMLCountry, int qbXMLMajorVers, int qbXMLMinorVers)
        {
            if (Session["counter"] == null)
            {
                Session["counter"] = 0;
            }

            string evLogTxt = string.Format(
@"WebMethod: sendRequestXML() has been called by QBWebconnector

Parameters received:
string ticket = {0}
string strHCPResponse = {1}
string strCompanyFileName = {2}
string qbXMLCountry = {3}
int qbXMLMajorVers = {4}
int qbXMLMinorVers = {5}

",
             ticket,
             strHCPResponse,
             strCompanyFileName,
             qbXMLCountry,
             qbXMLMajorVers.ToString(),
             qbXMLMinorVers.ToString() 
             );
             
            ArrayList req = BuildBusinessRequestMessage();
            string request = "";
            int total = req.Count;
            count = Convert.ToInt32(Session["counter"]);

            if (count < total)
            {
                request = req[count].ToString();
                evLogTxt = evLogTxt + "sending request no = " + (count + 1) + "\r\n";
                Session["counter"] = ((int)Session["counter"]) + 1;
            }
            else
            {
                count = 0;
                Session["counter"] = 0;
                request = "";
            }
            evLogTxt = evLogTxt + "\r\n";
            evLogTxt = evLogTxt + "Return values: " + "\r\n";
            evLogTxt = evLogTxt + "string request = " + request + "\r\n";
            LogEvent(evLogTxt);
            return request;
        }


        /// <summary>
        /// WebMethod# 5 - receiveResponseXML()
        /// Signature: public int receiveResponseXML(string ticket, string response, string hresult, string message)
        /// 
        /// IN: 
        /// string ticket
        /// string response
        /// string hresult
        /// string message
        ///
        /// OUT: 
        /// int retVal
        /// Greater than zero  = There are more request to send
        /// 100 = Done. no more request to send
        /// Less than zero  = Custom Error codes
        /// </summary>
        [WebMethod(Description = "This web method facilitates web service to receive response XML from QuickBooks via QBWebConnector", 
            EnableSession = true)]
        public int receiveResponseXML(string ticket, string response, string hresult, string message)
        {
 
            var evLogTxt = string.Format(
@"WebMethod: receiveResponseXML() has been called by QBWebconnector

Parameters received:
string ticket = {0}
string response = {1}
string hresult = {2}
string message = {3}

",
             ticket,
             response,
             hresult,
             message 
            ); 

            int retVal = 0;
            if (!hresult.ToString().Equals(""))
            {
                // if there is an error with response received, web service could also return a -ve int		
                evLogTxt = evLogTxt + "HRESULT = " + hresult + "\r\n";
                evLogTxt = evLogTxt + "Message = " + message + "\r\n";
                retVal = -101;
            }
            else
            {
                evLogTxt = evLogTxt + "Length of response received = " + response.Length + "\r\n";

                ArrayList req = BuildBusinessRequestMessage();
                int total = req.Count;
                int count = Convert.ToInt32(Session["counter"]);

                int percentage = (count * 100) / total;
                if (percentage >= 100)
                {
                    count = 0;
                    Session["counter"] = 0;
                }
                retVal = percentage;
            }
            evLogTxt = evLogTxt + "\r\n";
            evLogTxt = evLogTxt + "Return values: " + "\r\n";
            evLogTxt = evLogTxt + "int retVal= " + retVal.ToString() + "\r\n";
            LogEvent(evLogTxt);
            return retVal;
        }


        /// <summary>
        /// Build QBXML Business Request Message
        /// </summary>
        /// <returns></returns>
        public ArrayList BuildBusinessRequestMessage()
        { 
            try
            {

                var connectionString = BuildMongoConnectionString();

                var client = new MongoClient(connectionString);
                var mongo = client.GetServer();
                mongo.Connect(TimeSpan.FromSeconds(Settings.Default.dbhostconnecttimeout));

                //var credentials = new MongoCredentials(Settings.Default.appusername, Settings.Default.appuserpassword);
                var db = mongo.GetDatabase(Settings.Default.dbcatalogname
                    //,credentials
                    );
                 
                using (mongo.RequestStart(db))
                {
                    var collection = db.GetCollection<BsonDocument>("contacts");

                    foreach (BsonDocument contact in collection.FindAll())
                    {
                        var xml = new XDocument(
                          new XDeclaration("1.0", "UTF-8", null),
                          new XElement("Orders",
                              new XElement("QBXML",
                                  new XElement("QBXMLMsgsRq",
                                      new XAttribute("onError", "stopOnError"),
                                      new XElement("CustomerAddRq",
                                          new XAttribute("requestID", "2"), 
                                                  new XElement("custAdd",
                                                      new XElement("Name", contact.GetValue("companyName") ?? string.Empty),
                                                      new XElement("FirstName", contact.GetValue("firstName") ?? string.Empty),
                                                      new XElement("LastName", contact.GetValue("lastName") ?? string.Empty),
                                                      new XElement("Email", contact.GetValue("email") ?? string.Empty))
                                          )
                                      )
                                  )
                              ) 
                          ); 
                        xml.AddFirst(new XProcessingInstruction("qbxml", "version=\"2.0\""));
                         
                        req.Add(SerializeXDoc(xml));
                            
                    } 
                }

            }
            catch (Exception exc)
            {
                //mongo connect fail
                LogEvent(exc.StackTrace);
            }

            return req;
        }

         

        /// <summary>
        /// WebMethod - getLastError()
        /// Signature: public string getLastError(string ticket)
        /// 
        /// IN:
        /// string ticket
        /// 
        /// OUT:
        /// string retVal
        /// Possible Values:
        /// Error message describing last web service error
        /// </summary>
        [WebMethod]
        public string getLastError(string ticket)
        {
            string evLogTxt = "WebMethod: getLastError() has been called by QBWebconnector" + "\r\n\r\n";
            evLogTxt = evLogTxt + "Parameters received:\r\n";
            evLogTxt = evLogTxt + "string ticket = " + ticket + "\r\n";
            evLogTxt = evLogTxt + "\r\n";

            int errorCode = 0;
            string retVal = null;
            if (errorCode == -101)
            {
                retVal = "QuickBooks was not running!"; // This is just an example of custom user errors
            }
            else
            {
                retVal = "Error!";
            }
            evLogTxt = evLogTxt + "\r\n";
            evLogTxt = evLogTxt + "Return values: " + "\r\n";
            evLogTxt = evLogTxt + "string retVal= " + retVal + "\r\n";
            LogEvent(evLogTxt);
            return retVal;
        }


        /// <summary>
        /// WebMethod - closeConnection()
        /// At the end of a successful update session, QBWebConnector will call this web method.
        /// Signature: public string closeConnection(string ticket)
        /// 
        /// IN:
        /// string ticket 
        /// 
        /// OUT:
        /// string closeConnection result 
        /// </summary>
        [WebMethod]
        public string closeConnection(string ticket)
        {
            string evLogTxt = "WebMethod: closeConnection() has been called by QBWebconnector" + "\r\n\r\n";
            evLogTxt = evLogTxt + "Parameters received:\r\n";
            evLogTxt = evLogTxt + "string ticket = " + ticket + "\r\n";
            evLogTxt = evLogTxt + "\r\n";
            string retVal = null;

            retVal = "OK";

            evLogTxt = evLogTxt + "\r\n";
            evLogTxt = evLogTxt + "Return values: " + "\r\n";
            evLogTxt = evLogTxt + "string retVal= " + retVal + "\r\n";
            LogEvent(evLogTxt);
            return retVal;
        }




        #region Utility section


        private void LogEvent(string logText)
        {
            try
            {
                evLog.WriteEntry(logText);
            }
            catch { };
            return;
        }

        private static string SerializeXDoc(XDocument xml)
        {
            var s_builder = new StringBuilder();
            using (var writer = new StringWriter(s_builder))
            {
                xml.Save(writer);
            }
            return s_builder.ToString();
        }

        private static string BuildMongoConnectionString()
        {
            var dbhostport = string.Empty;
            if (Settings.Default.dbhostport != string.Empty)
            {
                if (Settings.Default.dbhostaddress.IndexOf(":") < 0)
                {
                    dbhostport = string.Format(":{0}", Settings.Default.dbhostport ?? "27017");
                }
            }

            var connectionString = string.Format("mongodb://{0}{1}", Settings.Default.dbhostaddress, dbhostport);
            return connectionString;
        }


        #endregion
    }
}
