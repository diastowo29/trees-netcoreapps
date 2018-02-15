using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Authenticators;
using System.Net.Mail;

using System.Net.Http;
using System.Net;

namespace newConsole {
  class CallingApi {
    string zendeskDomain = "https://developmenttesting.zendesk.com";
    string zendeskUsername = "rahdityoluhung89@gmail.com";
    string originDirectory = "D:/WORK/Doc/";
    string destDirectory = "/home/diastowo/Documents/DOT NET/excel done/";
    string zendeskPassword = "W3lcome123";

    public string callApiPost (object parameterBody, string url) {
      Console.WriteLine("CALL POST: " + url);

      var client = new RestClient(zendeskDomain + url);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.POST);
      // Console.WriteLine(JsonConvert.SerializeObject(parameterBody));
      request.AddParameter("application/json", parameterBody, ParameterType.RequestBody);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      return content;
    }

    public string callApi (String urls) {
      Console.WriteLine("CALL GET: " + urls);

      var client = new RestClient(zendeskDomain + urls);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.GET);
      // easily add HTTP Headers
      // request.AddHeader("Authorization", "Basic " + zendeskToken);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      return content;
    }
	}
}