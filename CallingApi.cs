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
using System.Net;

namespace newConsole {
  class CallingApi {
    // string zendeskDomain = "https://fifgroup1481257374.zendesk.com";
    string zendeskDomain = "https://fifgroup.zendesk.com";
    // string zendeskUsername = "eldien.hasmanto@treessolutions.com";
    // string zendeskPassword = "W3lcome123";

    public string callApiPost (object parameterBody, string url, string zendeskUsername, string zendeskPassword) {
      Console.WriteLine("CALL POST: " + url);
      Console.WriteLine(parameterBody);

      var client = new RestClient(zendeskDomain + url);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.POST);
      // Console.WriteLine(JsonConvert.SerializeObject(parameterBody));
      request.AddParameter("application/json", parameterBody, ParameterType.RequestBody);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      Console.WriteLine(response.StatusCode);
      if (response.StatusCode.ToString() == "Unauthorized" || response.StatusCode.ToString() == "0") {
        Console.WriteLine("Wrong Username or Password.. Program will exit");
        Environment.Exit(0);
      }
      return content;
    }

    public string callApi (String urls, string zendeskUsername, string zendeskPassword) {
      Console.WriteLine("CALL GET: " + urls);

      var client = new RestClient(zendeskDomain + urls);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.GET);
      // easily add HTTP Headers
      // request.AddHeader("Authorization", "Basic " + zendeskToken);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      Console.WriteLine(response.StatusCode);
      if (response.StatusCode.ToString() == "Unauthorized" || response.StatusCode.ToString() == "0") {
        Console.WriteLine("Wrong Username or Password.. Program will exit");
        Environment.Exit(0);
      }
      return content;
    }

    public string callApiPut (object parameterBody, string url, string zendeskUsername, string zendeskPassword) {
      Console.WriteLine("CALL PUT: " + url);
      Console.WriteLine(parameterBody);

      var client = new RestClient(zendeskDomain + url);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.PUT);
      // Console.WriteLine(JsonConvert.SerializeObject(parameterBody));
      request.AddParameter("application/json", parameterBody, ParameterType.RequestBody);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      Console.WriteLine(response.StatusCode);
      if (response.StatusCode.ToString() == "Unauthorized" || response.StatusCode.ToString() == "0") {
        Console.WriteLine("Wrong Username or Password.. Program will exit");
        Environment.Exit(0);
      }
      return content;
    }

    public string callApiDelete (string url, string zendeskUsername, string zendeskPassword) {
      Console.WriteLine("CALL DELETE: " + url);

      var client = new RestClient(zendeskDomain + url);
      client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

      var request = new RestRequest("", Method.DELETE);

      IRestResponse response = client.Execute(request);
      var content = response.Content;
      Console.WriteLine(response.StatusCode);
      if (response.StatusCode.ToString() == "Unauthorized" || response.StatusCode.ToString() == "0") {
        Console.WriteLine("Wrong Username or Password.. Program will exit");
        Environment.Exit(0);
      }
      return content;
    }
	}
}