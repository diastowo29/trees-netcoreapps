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
	class Organizations {
		public string searchOrganizations (string orgName, string groupMm, string userMm, string groupDh, string userDh, string groupBm, string userBm, string groupAh, string userAh) {
			CallingApi callingApi = new CallingApi();
			string orgId = "";
            string srcOrganization = "/api/v2/search.json?query=type:organization%20name:\"" + orgName + "\"";
            string srcOrgResponse = callingApi.callApi(srcOrganization);
            JObject srcOrgJoResponse = JObject.Parse(srcOrgResponse);
            JArray srcOrgArray = (JArray)srcOrgJoResponse["results"];
            if (srcOrgArray.Count == 0) {
                orgId = doCreateOrganizations(orgName, groupMm, userMm, groupDh, userDh, groupBm, userBm, groupAh, userAh);
            } else {
                orgId = srcOrgArray[0]["id"].ToString();
            	doUpdateOrganizations(orgId, groupMm, userMm, groupDh, userDh, groupBm, userBm, groupAh, userAh);
            }
            return orgId;
		}

		public string doCreateOrganizations (string orgName,string groupMm, string userMm, string groupDh, string userDh, string groupBm, string userBm, string groupAh, string userAh) {
			CallingApi callingApi = new CallingApi();
			Console.WriteLine("doCreateOrganizations");
			Dictionary<string,string> orgCustomField = new Dictionary<string,string>();
			orgCustomField.Add("group_mm", groupMm);
			orgCustomField.Add("user_mm", userMm);
			orgCustomField.Add("group_dh", groupDh);
			orgCustomField.Add("user_dh", userDh);
			orgCustomField.Add("group_bm", groupBm);
			orgCustomField.Add("user_bm", userBm);
			orgCustomField.Add("group_ah", groupAh);
			orgCustomField.Add("user_ah", userAh);
			Dictionary<string,object> orgField = new Dictionary<string,object>();
			orgField.Add("organization_fields", orgCustomField);
			orgField.Add("name", orgName);
			Dictionary<string,object> org = new Dictionary<string,object>();
			org.Add("organization", orgField);
            Console.WriteLine(JsonConvert.SerializeObject(org));

            string orgId = "";
            string createOrganization = "/api/v2/organizations.json";
            string createOrgResponse = callingApi.callApiPost(JsonConvert.SerializeObject(org), createOrganization);
            JObject createOrgJoResponse = JObject.Parse(createOrgResponse);
            JObject createOrg = (JObject)createOrgJoResponse["organization"];
            orgId = createOrg["id"].ToString();

            return orgId;
		}

		public void doUpdateOrganizations (string orgId, string groupMm, string userMm, string groupDh, string userDh, string groupBm, string userBm, string groupAh, string userAh) {
			Console.WriteLine("doUpdateOrganizations");
		}
	}
}