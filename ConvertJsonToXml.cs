using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace JsonToXml
{
    public class Version
    {
        public static string ConvertJsonToXml(string jsonUrl)
        {
            // Download the JSON content from the URL
            string jsonContent;
            ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (WebClient client = new WebClient())
            {
                jsonContent = client.DownloadString(jsonUrl);
            }
            // Parse the JSON content
            JObject jsonObject = JObject.Parse(jsonContent);

            // Get the root element name
            string rootElementName = jsonObject.Properties().FirstOrDefault()?.Name;
            if (string.IsNullOrEmpty(rootElementName))
            {
                throw new Exception("Unable to determine the root element name in the JSON.");
            }
            // Convert the JSON to XML
            XmlDocument xmlDoc = JsonConvert.DeserializeXmlNode(jsonContent, rootElementName);

            // Convert the XmlDocument to XDocument
            XDocument xDoc;
            using (MemoryStream stream = new MemoryStream())
            {
                xmlDoc.Save(stream);
                stream.Position = 0;
                xDoc = XDocument.Load(stream);
            }

            // Return the XML string
            return xDoc.ToString();
        }
        /// <summary>
        /// Update Json File After Update New Version
        /// </summary>
        /// <param name="json"></param>
        public static void save(string json)
        {
            File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "/checker.json", json);
        }
    }
}
