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

namespace ArdaQBWS
{
    /// <summary>
    /// Summary description for QBWebService
    /// </summary>
    [WebService(
         Namespace = "http://developer.intuit.com/",
         Name = "WCWebService",
         Description = "Sample WebService in ASP.NET to demonstrate " +
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

            // For simplicity of sample, a hardcoded username/password is used.
            string pwd = "password";

            if (strUserName.Trim().Equals("test") && strPassword.Trim().Equals(pwd))
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

        [WebMethod(Description = "This web method facilitates web service to send request XML to QuickBooks via QBWebConnector", EnableSession = true)]
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
        public string sendRequestXML(string ticket, string strHCPResponse, string strCompanyFileName,
            string qbXMLCountry, int qbXMLMajorVers, int qbXMLMinorVers)
        {
            if (Session["counter"] == null)
            {
                Session["counter"] = 0;
            }
            string evLogTxt = "WebMethod: sendRequestXML() has been called by QBWebconnector" + "\r\n\r\n";
            evLogTxt = evLogTxt + "Parameters received:\r\n";
            evLogTxt = evLogTxt + "string ticket = " + ticket + "\r\n";
            evLogTxt = evLogTxt + "string strHCPResponse = " + strHCPResponse + "\r\n";
            evLogTxt = evLogTxt + "string strCompanyFileName = " + strCompanyFileName + "\r\n";
            evLogTxt = evLogTxt + "string qbXMLCountry = " + qbXMLCountry + "\r\n";
            evLogTxt = evLogTxt + "int qbXMLMajorVers = " + qbXMLMajorVers.ToString() + "\r\n";
            evLogTxt = evLogTxt + "int qbXMLMinorVers = " + qbXMLMinorVers.ToString() + "\r\n";
            evLogTxt = evLogTxt + "\r\n";

            ArrayList req = buildRequest();
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
            logEvent(evLogTxt);
            return request;
        }


        [WebMethod(Description = "This web method facilitates web service to receive response XML from QuickBooks via QBWebConnector", EnableSession = true)]
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
        public int receiveResponseXML(string ticket, string response, string hresult, string message)
        {
            string evLogTxt = "WebMethod: receiveResponseXML() has been called by QBWebconnector" + "\r\n\r\n";
            evLogTxt = evLogTxt + "Parameters received:\r\n";
            evLogTxt = evLogTxt + "string ticket = " + ticket + "\r\n";
            evLogTxt = evLogTxt + "string response = " + response + "\r\n";
            evLogTxt = evLogTxt + "string hresult = " + hresult + "\r\n";
            evLogTxt = evLogTxt + "string message = " + message + "\r\n";
            evLogTxt = evLogTxt + "\r\n";

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

                ArrayList req = buildRequest();
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
            logEvent(evLogTxt);
            return retVal;
        }

        private void logEvent(string logText)
        {
            try
            {
                evLog.WriteEntry(logText);
            }
            catch { };
            return;
        }

        public ArrayList buildRequest()
        {
            string strRequestXML = "";
            //step2: create the qbXML request
            string[,] namesOfPeople = { { "Assemlio SSL", "Vullkan", "Halili", "vullkan@opendatakosovo.org", "Mr" }, 
                                        { "Open Data Kosovo", "Feride", "Adili", "feride@opendatakosovo.org", "Ms"},
                                        { "LunaKos", "Egzontina", "Krasniqi", "egzontina@opendatakosovo.org", "Ms"},
                                        { "Deliotte", "Hamdi", "Aliu", "hamdi@opendatakosovo.org", "Mr"},
                                        { "ARDA", "Alush", "Gashi", "alush@opendatakosovo.org" , "Mr"},
                                        { "E-Prokurimi", "Arben", "Maloku", "arbenM@opendatakosovo.org", "Mr"},
                                        { "E-Nsh Besa", "Vlora", "Thaqi", "vlora@opendatakosovo.org", "Ms"},
                                        { "E-Nsh Nobeli", "Rita", "Ora", "rita@opendatakosovo.org", "Ms"},
                                        { "CFLI", "Arben", "Avdiu", "arben@opendatakosovo.org", "Mr"},
                                        { "Arbeni-Tours", "Latif", "Latifi", "latif@opendatakosovo.org", "Mr"},
                                        { "Besa-Tours", "Besa", "Hoxha", "besa@opendatakosovo.org", "Ms"},
                                        { "Fadili-Tours", "Fadil", "Hamza", "fadil@opendatakosovo.org", "Mr"} };

            for (int i = 0; i < namesOfPeople.GetLength(0); i++)
            {
                XmlDocument inputXMLDoc = new XmlDocument();
                inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
                inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"2.0\""));
                XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
                inputXMLDoc.AppendChild(qbXML);
                XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
                qbXML.AppendChild(qbXMLMsgsRq);
                qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
                XmlElement custAddRq = inputXMLDoc.CreateElement("CustomerAddRq");
                qbXMLMsgsRq.AppendChild(custAddRq);
                custAddRq.SetAttribute("requestID", "2");
                XmlElement custAdd = inputXMLDoc.CreateElement("CustomerAdd");
                custAddRq.AppendChild(custAdd);
                custAdd.AppendChild(inputXMLDoc.CreateElement("Name")).InnerText = namesOfPeople[i, 0];
                custAdd.AppendChild(inputXMLDoc.CreateElement("CompanyName")).InnerText = namesOfPeople[i, 0];
                custAdd.AppendChild(inputXMLDoc.CreateElement("FirstName")).InnerText = namesOfPeople[i, 1];
                custAdd.AppendChild(inputXMLDoc.CreateElement("LastName")).InnerText = namesOfPeople[i, 2];
                custAdd.AppendChild(inputXMLDoc.CreateElement("Email")).InnerText = namesOfPeople[i, 3];

                strRequestXML = inputXMLDoc.OuterXml;
                req.Add(strRequestXML);

                // Clean up
                strRequestXML = "";
                inputXMLDoc = null;
                qbXMLMsgsRq = null;
                custAdd = null;
            }
            return req;
        }


        [WebMethod]
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
            logEvent(evLogTxt);
            return retVal;
        }


        [WebMethod]
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
            logEvent(evLogTxt);
            return retVal;
        }
    }
}
