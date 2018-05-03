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
	class Organizations {
		public string searchOrganizations (string orgName, string groupMm, string userMm, string groupDh, string userDh, string groupAh, string userAh, string zendeskUsername, string zendeskPassword) {
			CallingApi callingApi = new CallingApi();
			string orgId = "";
            string srcOrganization = "/api/v2/search.json?query=type:organization%20name:\"" + orgName + "\"";
            string srcOrgResponse = callingApi.callApi(srcOrganization, zendeskUsername, zendeskPassword);
            JObject srcOrgJoResponse = JObject.Parse(srcOrgResponse);
            JArray srcOrgArray = (JArray)srcOrgJoResponse["results"];
            if (srcOrgArray.Count == 0) {
                orgId = doCreateOrganizations(orgName, groupMm, userMm, groupDh, userDh, groupAh, userAh, zendeskUsername, zendeskPassword);
            } else {
                orgId = srcOrgArray[0]["id"].ToString();
            	doUpdateOrganizations(orgId, groupMm, userMm, groupDh, userDh, groupAh, userAh, zendeskUsername, zendeskPassword);
            }
            return orgId;
		}

		public string doCreateOrganizations (string orgName,string groupMm, string userMm, string groupDh, string userDh, string groupAh, string userAh, string zendeskUsername, string zendeskPassword) {
			CallingApi callingApi = new CallingApi();
			Console.WriteLine("doCreateOrganizations");
			Dictionary<string,string> orgCustomField = new Dictionary<string,string>();
			orgCustomField.Add("id_group_mm", groupMm);
			orgCustomField.Add("mm_id", userMm);
			orgCustomField.Add("id_group_dept_head", groupDh);
			orgCustomField.Add("dept_head_id", userDh);
			orgCustomField.Add("id_group_area_head", groupAh);
			orgCustomField.Add("area_head_id", userAh);
			Dictionary<string,object> orgField = new Dictionary<string,object>();
			orgField.Add("organization_fields", orgCustomField);
			orgField.Add("name", orgName);
			Dictionary<string,object> org = new Dictionary<string,object>();
			org.Add("organization", orgField);

            string orgId = "";
            string createOrganization = "/api/v2/organizations.json";
            string createOrgResponse = callingApi.callApiPost(JsonConvert.SerializeObject(org), createOrganization, zendeskUsername, zendeskPassword);
            JObject createOrgJoResponse = JObject.Parse(createOrgResponse);
            JObject createOrg = (JObject)createOrgJoResponse["organization"];
            orgId = createOrg["id"].ToString();

            return orgId;
		}

		public string doUpdateOrganizations (string orgId, string groupMm, string userMm, string groupDh, string userDh, string groupAh, string userAh, string zendeskUsername, string zendeskPassword) {
			Console.WriteLine("doUpdateOrganizations");
			CallingApi callingApi = new CallingApi();
			Dictionary<string,string> orgCustomField = new Dictionary<string,string>();
			orgCustomField.Add("id_group_mm", groupMm);
			orgCustomField.Add("mm_id", userMm);
			orgCustomField.Add("id_group_dept_head", groupDh);
			orgCustomField.Add("dept_head_id", userDh);
			orgCustomField.Add("id_group_area_head", groupAh);
			orgCustomField.Add("area_head_id", userAh);
			Dictionary<string,object> orgField = new Dictionary<string,object>();
			orgField.Add("organization_fields", orgCustomField);
			// orgField.Add("name", orgName);
			Dictionary<string,object> org = new Dictionary<string,object>();
			org.Add("organization", orgField);

            string orgIdUpdate = "";
            string updateOrganization = "/api/v2/organizations/" + orgId + ".json";
            string updateOrgResponse = callingApi.callApiPut(JsonConvert.SerializeObject(org), updateOrganization, zendeskUsername, zendeskPassword);
            JObject updateOrgJoResponse = JObject.Parse(updateOrgResponse);
            JObject updateOrg = (JObject)updateOrgJoResponse["organization"];
            orgIdUpdate = updateOrg["id"].ToString();

            return orgIdUpdate;
		}
	}
}