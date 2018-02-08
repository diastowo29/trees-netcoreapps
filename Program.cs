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

namespace newConsole
{
    class Program
    {
        static string zendeskDomain = "https://developmenttesting.zendesk.com";
        static string zendeskUsername = "rahdityoluhung89@gmail.com";
        static string destDirectory = "/home/diastowo/Documents/DOT NET/excel done/";
        static List<object> userList = new List<object>();
        static string supportGroupId = "";
        static List<Dictionary<string,string>> doneList = new List<Dictionary<string,string>>();

        // static string zendeskDomain = "https://treesdemo1.zendesk.com";
        // static string zendeskUsername = "eldien.hasmanto@treessolutions.com";

        static string zendeskPassword = "W3lcome123";

        static List<JToken> allGroups = new List<JToken>();
        
        static void Main(string[] args)
        {
            // // // // // deleteGroups();
            initiate();
        	System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        	string[] array1 = Directory.GetFiles(@"/home/diastowo/Documents/DOT NET/excel/");
        	foreach (string filePath in array1) {
        		if (filePath.Contains("xlsx")) {
        			doXlsx(filePath);
        		} else if (filePath.Contains("xls")) {
        			// doXls(filePath);
        		}
                // File.Copy(filePath, destDirectory + Path.GetFileName(filePath));
        	}

        }

        public static void doXlsx (string filePath) {
            Dictionary<string, string> mappingList = new Dictionary<string, string>();
        	Console.WriteLine("===== DO XLSX =====");
            List<String> keys = new List<String>();
            List<Dictionary<string, string>> mappingArray = new List<Dictionary<string, string>>();
            int skipIndex = 0;

        	Console.WriteLine(filePath);
        	var package = new ExcelPackage(new FileInfo(filePath));
			ExcelWorksheet sheet = package.Workbook.Worksheets[0];
			var rowCount = sheet.Dimension.End.Row;
			var colCount = sheet.Dimension.End.Column;
			for (int i=1; i<=rowCount; i++) {
                mappingList = new Dictionary<string, string>();
				// Console.WriteLine("===== NEW ROW =====");
                int myColCounter = 0;

				for (int j=1; j<=colCount; j++) {
                    if (i == 1) {
                        bool keyExist = false;
                        for (int k=0; k<keys.Count; k++) {
                            if (keys[k] == sheet.Cells[i,j].Value.ToString()) {
                                skipIndex = j;
                            }
                        }
                        keys.Add(sheet.Cells[i,j].Value.ToString());
                    } else {
                        if (j != skipIndex) {
                            string values = "";
                            if (sheet.Cells[i,j].Value == null) {
                                values = "";
                            } else {
                                values = sheet.Cells[i,j].Value.ToString();
                            }
                            mappingList.Add(keys[j-1], values);
                            myColCounter++;
                        }
                    }
				}
                if (i != 1) {
                    mappingArray.Add(mappingList);
                }
			}
            doProcessMapping(mappingArray);
            // string jsonString = JsonConvert.SerializeObject(mappingArray);
            // Console.WriteLine(jsonString);
        }

        public static void doXls (string filePath) {
            Dictionary<string, string> mappingList = new Dictionary<string, string>();
            List<Dictionary<string, string>> mappingArray = new List<Dictionary<string, string>>();
            var rowCount = 0;
            List<String> keys = new List<String>();
            int skipIndex = 0;

        	Console.WriteLine("===== DO XLS =====");
        	Console.WriteLine(filePath);
        	using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read)) {
				using (var reader = ExcelReaderFactory.CreateReader(stream)) {
					do {
						while (reader.Read()) {
                            rowCount++;
                            mappingList = new Dictionary<string, string>();
							// Console.WriteLine("===== NEW ROW =====");
							for (int i=0; i<reader.FieldCount; i++) {
                                if (rowCount == 1) {
                                    for (int k=0; k<keys.Count; k++) {
                                        if (keys[k] == reader.GetValue(i).ToString()) {
                                            skipIndex = i;
                                        }
                                    }
                                    keys.Add(reader.GetValue(i).ToString());
                                } else {
                                    if (i != skipIndex) {
                                        string values = "";
                                        if (reader.GetValue(i) == null) {
                                            values = "";
                                        } else {
                                            values = reader.GetValue(i).ToString();
                                        }
                                        mappingList.Add(keys[i], values);
                                    }
                                }
							}
                            string jsonString = JsonConvert.SerializeObject(mappingList);
                            mappingArray.Add(mappingList);
                            // Console.WriteLine(jsonString);
						}
					} while (reader.NextResult());

					// 2. Use the AsDataSet extension method
					// var result = reader.AsDataSet();

					// The result of each spreadsheet is in result.Tables
				}
			}
        }

        public static void doProcessMapping(List<Dictionary<string, string>> entries){
            string userMmId = "";
            string userDhId = "";
            string userBmId = "";
            string userAhId = "";

            string groupMMid = "";
            string groupBMid = "";
            string groupAHid = "";
            string groupDHid = "";
            doneList = new List<Dictionary<string,string>>();

            for (int i=0; i<entries.Count; i++) {
                string groupMM = "Group MM " + entries[i]["Branch Code"].ToString();
                string groupBM = "Group BM " + entries[i]["Branch Code"].ToString();
                string groupAH = "Group Area Head " + entries[i]["AREA"].ToString();
                string groupDH = "Group Dept Head " + entries[i]["Region Dept Head"].ToString();
                string orgName = "Organization " + entries[i]["Branch Code"].ToString();

                string namaMm = entries[i]["MM"].ToString();
                string namaDh = entries[i]["Dept. Head"].ToString();
                string namaBm = entries[i]["NAMA BM"].ToString();
                string namaAh = entries[i]["NAMA Area Head"].ToString();

                string dealerName = entries[i]["Dealer Name"].ToString();
                string dealerId = entries[i]["Dealer ID"].ToString();
                string emailPicOwner = entries[i]["Alamat email PIC Owner"].ToString();
                string emailOwner = entries[i]["Alamat email owner"].ToString();
                string emailPicOutlet = entries[i]["Alamat email PIC Outlet"].ToString();

                Program newCon = new Program();
                if (i==0) {
                    // callApi(entries[i]);

                    if (isGroupExist(groupMM) == "0") {
                        string createResponse = newCon.doCreateGroup(groupMM);
                        JObject createObject = JObject.Parse(createResponse);
                        JObject group = (JObject)createObject["group"];
                        groupMMid = group["id"].ToString();
                    } else {
                        groupMMid = isGroupExist(groupMM);
                    }
                    if (isGroupExist(groupBM) == "0") {
                        string createResponse = newCon.doCreateGroup(groupBM);
                        JObject createObject = JObject.Parse(createResponse);
                        JObject group = (JObject)createObject["group"];
                        groupBMid = group["id"].ToString();
                    } else {
                        groupBMid = isGroupExist(groupBM);
                    }
                    if (isGroupExist(groupAH) == "0") {
                        string createResponse = newCon.doCreateGroup(groupAH);
                        JObject createObject = JObject.Parse(createResponse);
                        JObject group = (JObject)createObject["group"];
                        groupAHid = group["id"].ToString();
                    } else {
                        groupAHid = isGroupExist(groupAH);
                    }
                    if (isGroupExist(groupDH) == "0") {
                        string createResponse = newCon.doCreateGroup(groupDH);
                        JObject createObject = JObject.Parse(createResponse);
                        JObject group = (JObject)createObject["group"];
                        groupDHid = group["id"].ToString();
                    } else {
                        groupDHid = isGroupExist(groupDH);
                    }

                    string nameParameter = "name:\"" + namaMm + "\"+name:\"" + namaDh + "\"+name:\"" + namaBm + "\"+name:\"" + namaAh + "\"";
                    string userResponse = isUserExist(nameParameter);
                    JObject userJoResponse = JObject.Parse(userResponse);
                    JArray usersList = (JArray)userJoResponse["results"];
                    bool mmFound = false;
                    bool dhFound = false;
                    bool bmFound = false;
                    bool ahFound = false;
                    for (int u=0; u<usersList.Count; u++){
                        if (usersList[u]["name"].ToString().ToLower() == namaMm.ToLower()) {
                            mmFound = true;
                            userMmId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user mm id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaDh.ToLower()) {
                            dhFound = true;
                            userDhId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user dh id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaBm.ToLower()) {
                            bmFound = true;
                            userBmId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user bm id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaAh.ToLower()) {
                            ahFound = true;
                            userAhId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user ah id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                    }

                    if (!mmFound) {
                        string userCreate = newCon.createUser(namaMm);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userMmId = user["id"].ToString();
                    }
                    if (!dhFound) {
                        string userCreate = newCon.createUser(namaDh);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userDhId = user["id"].ToString();
                    }
                    if (!bmFound) {
                        string userCreate = newCon.createUser(namaBm);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userBmId = user["id"].ToString();
                    }
                    if (!ahFound) {
                        string userCreate = newCon.createUser(namaAh);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userAhId = user["id"].ToString();
                    }

                    List<Dictionary<string,string>> groupMembershipsList = createGroupMembership(groupMMid, userMmId, groupDHid, userDhId, groupBMid, userBmId, groupAHid, userAhId);

                    // List<string> groupsIds = new List<string>();
                    // groupsIds.Add(groupMMid);
                    // groupsIds.Add(groupDHid);
                    // groupsIds.Add(groupBMid);
                    // groupsIds.Add(groupAHid);
                    checkGroupMemberships(groupMembershipsList);

                } else if (i <= 10) {

                    if (entries[i]["Branch Code"].ToString() != entries[i-1]["Branch Code"].ToString()) {
                        if (isGroupExist(groupMM) == "0") {
                            string createResponse = newCon.doCreateGroup(groupMM);
                            JObject createObject = JObject.Parse(createResponse);
                            JObject group = (JObject)createObject["group"];
                            groupMMid = group["id"].ToString();
                        } else {
                            groupMMid = isGroupExist(groupMM);
                        }

                        if (isGroupExist(groupBM) == "0") {
                            string createResponse = newCon.doCreateGroup(groupBM);
                            JObject createObject = JObject.Parse(createResponse);
                            JObject group = (JObject)createObject["group"];
                            groupBMid = group["id"].ToString();
                        } else {
                            groupBMid = isGroupExist(groupBM);
                        }
                    }

                    if (entries[i]["AREA"].ToString() != entries[i-1]["AREA"].ToString()) {
                        if (isGroupExist(groupAH) == "0") {
                            string createResponse = newCon.doCreateGroup(groupAH);
                            JObject createObject = JObject.Parse(createResponse);
                            JObject group = (JObject)createObject["group"];
                            groupAHid = group["id"].ToString();
                        } else {
                            groupAHid = isGroupExist(groupAH);
                        }
                    }

                    if (entries[i]["Region Dept Head"].ToString() != entries[i-1]["Region Dept Head"].ToString()) {
                        if (isGroupExist(groupDH) == "0") {
                            string createResponse = newCon.doCreateGroup(groupDH);
                            JObject createObject = JObject.Parse(createResponse);
                            JObject group = (JObject)createObject["group"];
                            groupDHid = group["id"].ToString();
                        } else {
                            groupDHid = isGroupExist(groupDH);
                        }
                    }


                    string nameParameter = "name:\"" + namaMm + "\"+name:\"" + namaDh + "\"+name:\"" + namaBm + "\"+name:\"" + namaAh + "\"";
                    string userResponse = isUserExist(nameParameter);
                    JObject userJoResponse = JObject.Parse(userResponse);
                    JArray usersList = (JArray)userJoResponse["results"];
                    bool mmFound = false;
                    bool dhFound = false;
                    bool bmFound = false;
                    bool ahFound = false;
                    for (int u=0; u<usersList.Count; u++){
                        if (usersList[u]["name"].ToString().ToLower() == namaMm.ToLower()) {
                            mmFound = true;
                            userMmId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user mm id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaDh.ToLower()) {
                            dhFound = true;
                            userDhId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user dh id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaBm.ToLower()) {
                            bmFound = true;
                            userBmId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user bm id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                        if (usersList[u]["name"].ToString().ToLower() == namaAh.ToLower()) {
                            ahFound = true;
                            userAhId = usersList[u]["id"].ToString();
                            // Console.WriteLine("get user ah id");
                            // Console.WriteLine(usersList[u]["id"]);
                        }
                    }

                    if (!mmFound) {
                        string userCreate = newCon.createUser(namaMm);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userMmId = user["id"].ToString();
                    }
                    if (!dhFound) {
                        string userCreate = newCon.createUser(namaDh);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userDhId = user["id"].ToString();
                    }
                    if (!bmFound) {
                        string userCreate = newCon.createUser(namaBm);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userBmId = user["id"].ToString();
                    }
                    if (!ahFound) {
                        string userCreate = newCon.createUser(namaAh);
                        JObject userCreateJoResponse = JObject.Parse(userCreate);
                        JObject user = (JObject)userJoResponse["user"];
                        userAhId = user["id"].ToString();
                    }

                    List<Dictionary<string,string>> groupMembershipsList = createGroupMembership(groupMMid, userMmId, groupDHid, userDhId, groupBMid, userBmId, groupAHid, userAhId);
                    // List<string> groupsIds = new List<string>();
                    // groupsIds.Add(groupMMid);
                    // groupsIds.Add(groupDHid);
                    // groupsIds.Add(groupBMid);
                    // groupsIds.Add(groupAHid);
                    checkGroupMemberships(groupMembershipsList);
                checkDealer(dealerId, dealerName, emailOwner, emailPicOwner, emailPicOutlet);
                }
            }
        }

        public static void checkDealer(string dealerId, string dealerName, string emailOwner, string emailPicOwner, string emailPicOutlet) {
            string createUpdateManyUser = zendeskDomain + "/api/v2/users/create_or_update_many.json";

            Dictionary<string, object> usersDict = new Dictionary<string,object>();
            Dictionary<string, string> userFields = new Dictionary<string,string>();
            Dictionary<string, object> newUser = new Dictionary<string,object>();
            if (emailOwner != "#N/A") {
                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "Owner");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailOwner);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailOwner);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            if (emailPicOwner != "#N/A") {
                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "PIC Owner");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailPicOwner);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailPicOwner);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            if (emailPicOutlet != "#N/A") {
                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "PIC Outlet");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailPicOutlet);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailPicOutlet);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            
            // userFields.Add("dealer_id", dealerId);
            // userFields.Add("email_owner", emailOwner);
            // userFields.Add("email_pic_owner", emailPicOwner);
            // userFields.Add("email_pic_outlet", emailPicOutlet);
            // newUser.Add("name", dealerName);
            // newUser.Add("user_fields", userFields);

            usersDict.Add("users", userList);
            Console.WriteLine(JsonConvert.SerializeObject(usersDict));

            // string createUpdateUser = callApiPost();
        }

        public static void checkGroupMemberships (List<Dictionary<string,string>> groupsIds) {

            /*MAKE IT IF REACH 100 ARRAY THEN EXECUTE*/
            List<string> willBeDelete = new List<string>();
            for (int i=0; i<groupsIds.Count; i++){
                string groupMembershipApi = zendeskDomain + "/api/v2/groups/" + groupsIds[i]["group_id"] + "/memberships.json";
                string groupMembershipRseponse = callApi(groupMembershipApi);
                // Console.WriteLine(groupMembershipRseponse);
                JObject groupMembershipJoResponse = JObject.Parse(groupMembershipRseponse);
                JArray memberList = (JArray)groupMembershipJoResponse["group_memberships"];

                if (memberList.Count > 1) {
                    for (int j=0; j<memberList.Count; j++) {
                        if (memberList[j]["user_id"].ToString() != groupsIds[i]["user_id"]) {
                            willBeDelete.Add(memberList[j]["id"].ToString());
                        }
                    }
                }

                string userGroupsApi = zendeskDomain + "/api/v2/users/" + groupsIds[i]["user_id"] + "/group_memberships.json";
                string userGroupResponse = callApi(userGroupsApi);
                JObject userGroupJoResponse = JObject.Parse(userGroupResponse);
                JArray groupsList = (JArray)userGroupJoResponse["group_memberships"];
                bool groupFound = false;
                if (groupsList.Count > 1) {
                    for (int k=0; k<groupsList.Count; k++) {
                        if (groupsList[k]["group_id"].ToString() != groupsIds[i]["group_id"]) {
                            if (groupsList[k]["group_id"].ToString() != supportGroupId) {
                                for (int l=0; l<doneList.Count; l++) {
                                    if (doneList[l]["user_id"].ToString() == groupsIds[i]["user_id"].ToString()) {
                                        if (doneList[l]["group_id"].ToString() == groupsList[k]["group_id"].ToString()){
                                            groupFound = true;
                                        }
                                    }
                                }
                                if (!groupFound) {
                                    willBeDelete.Add(groupsList[k]["id"].ToString());
                                }
                            }
                        }
                    }
                }
            }
            doDeleteMemberships(willBeDelete);
        }

        public static void doDeleteMemberships(List<string> memberIds) {
            if (memberIds.Count > 0) {
                StringBuilder deleteMembershipApi = new StringBuilder();
                deleteMembershipApi.Append(zendeskDomain);
                deleteMembershipApi.Append("/api/v2/group_memberships/destroy_many.json?ids=");

                for (int i=0; i<memberIds.Count; i++) {
                    deleteMembershipApi.Append(memberIds[i]);
                    if (i != memberIds.Count-1) {
                        deleteMembershipApi.Append(",");
                    }
                }

                Console.WriteLine("CALL DELETE: " + deleteMembershipApi);
                var client = new RestClient(deleteMembershipApi.ToString());
                client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

                var request = new RestRequest("", Method.DELETE);

                IRestResponse response = client.Execute(request);
                var content = response.Content;
            }
        }

        public static string isGroupExist (string groupName) {
            string groupFound = "0";
            for (int i=0; i<allGroups.Count; i++) {
                JObject group = (JObject)allGroups[i];
                if (group["name"].ToString() == groupName) {
                    groupFound = group["id"].ToString();
                }
            }
            return groupFound;
        }

        public static string isUserExist (string nameParameter) {
            var searchUserApi =  zendeskDomain + "/api/v2/search.json?query=type:user%20" + nameParameter;
            string response = callApi(searchUserApi);
            return response;
        }

        public string createUser (string userName) {
            Dictionary<string,object> newUser = new Dictionary<string,object>();
            Dictionary<string,string> userProp = new Dictionary<string,string>();
            userProp.Add("name", userName);
            newUser.Add("user", userProp);
            var createUserApi =  zendeskDomain + "/api/v2/users.json";
            Console.WriteLine(JsonConvert.SerializeObject(newUser));
            // string response = callApiPost(JsonConvert.SerializeObject(user), createUserApi);
            return "response";
        }

        public static List<Dictionary<string,string>> createGroupMembership (string groupMm, string userMm, string groupDh, string userDh, string groupBm, string userBm, string groupAh, string userAh) {
            var createGroupMembershipApi = zendeskDomain + "/api/v2/group_memberships/create_many.json";

            Dictionary<string,string> groupMembers = new Dictionary<string,string>();
            List<Dictionary<string,string>> groupMembersList = new List<Dictionary<string,string>>();
            Dictionary<string, List<Dictionary<string,string>>> groupMemberships = new Dictionary<string, List<Dictionary<string,string>>>();
            groupMembers.Add("user_id", userMm);
            groupMembers.Add("group_id", groupMm);
            doneList.Add(groupMembers);
            groupMembersList.Add(groupMembers);
            groupMembers = new Dictionary<string,string>();
            groupMembers.Add("user_id", userDh);
            groupMembers.Add("group_id", groupDh);
            doneList.Add(groupMembers);
            groupMembersList.Add(groupMembers);
            groupMembers = new Dictionary<string,string>();
            groupMembers.Add("user_id", userBm);
            groupMembers.Add("group_id", groupBm);
            doneList.Add(groupMembers);
            groupMembersList.Add(groupMembers);
            groupMembers = new Dictionary<string,string>();
            groupMembers.Add("user_id", userAh);
            groupMembers.Add("group_id", groupAh);
            doneList.Add(groupMembers);
            groupMembersList.Add(groupMembers);
            groupMembers = new Dictionary<string,string>();

            groupMemberships.Add("group_memberships", groupMembersList);
            Console.WriteLine(JsonConvert.SerializeObject(groupMemberships));
            string createMembershipResponse =  callApiPost(JsonConvert.SerializeObject(groupMemberships), createGroupMembershipApi);
            return groupMembersList;
        }

        public string doCreateGroup(string groupName){
            Dictionary<string,string> groupJson = new Dictionary<string,string>();
            Dictionary<string,Dictionary<string,string>> groupParameter = new Dictionary<string,Dictionary<string,string>>();

            groupJson.Add("name", groupName);
            groupParameter.Add("group", groupJson);
            var createGroupAPI =  zendeskDomain + "/api/v2/groups.json";
            Console.WriteLine(JsonConvert.SerializeObject(groupParameter));
            string content = callApiPost(JsonConvert.SerializeObject(groupParameter), createGroupAPI);
            return content;
        }

        /*public static string callApiPut (object parameterBody, string url) {
            Console.WriteLine("CALL PUT: " + url);

            var client = new RestClient(url);
            client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

            var request = new RestRequest("", Method.POST);
            // Console.WriteLine(JsonConvert.SerializeObject(parameterBody));
            request.AddParameter("application/json", parameterBody, ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var content = response.Content;
            return content;
        }*/

        public static string callApiPost (object parameterBody, string url) {
            Console.WriteLine("CALL POST: " + url);

            var client = new RestClient(url);
            client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

            var request = new RestRequest("", Method.POST);
            // Console.WriteLine(JsonConvert.SerializeObject(parameterBody));
            request.AddParameter("application/json", parameterBody, ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var content = response.Content;
            return content;
        }

        public static string callApi (String urls) {
            Console.WriteLine("CALL GET: " + urls);

            var client = new RestClient(urls);
            client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

            var request = new RestRequest("", Method.GET);
            // easily add HTTP Headers
            // request.AddHeader("Authorization", "Basic " + zendeskToken);

            IRestResponse response = client.Execute(request);
            var content = response.Content;
            return content;
        }

        public static void initiate () {
            getAllGroups("null");
        }

        public static void getAllGroups (string nextPage) {
            var getGroupApi =  zendeskDomain + "/api/v2/groups.json";
            string response = "";
            if (nextPage == "null") {
                response = callApi(getGroupApi);
            } else {
                response = callApi(nextPage);
            }

            JObject joResponse = JObject.Parse(response);
            JValue nextPageUrl = (JValue)joResponse["next_page"];
            JArray groupsList = (JArray)joResponse["groups"];

            for (int i=0; i<groupsList.Count; i++) {
                if (groupsList[i]["name"].ToString() == "Support") {
                    supportGroupId = groupsList[i]["id"].ToString();
                }
                allGroups.Add(groupsList[i]);
            }

            if (joResponse["next_page"].ToString() != String.Empty) {
                getAllGroups(joResponse["next_page"].ToString());
            }
        }

        public static void deleteGroups() {
            // for (int i=0; i<allGroups.Count; i++) {
            //     if (allGroups[i]["name"].ToString() == "Support") {
            //         Console.WriteLine(allGroups[i]);
            //     } else {
            //         string deleteGroupApi = zendeskDomain + "/api/v2/groups/" + allGroups[i]["id"] + ".json";
            //         Console.WriteLine("CALL DELETE: " + deleteGroupApi);

            //         var client = new RestClient(deleteGroupApi);
            //         client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

            //         var request = new RestRequest("", Method.DELETE);

            //         IRestResponse response = client.Execute(request);
            //         var content = response.Content;
            //         // return content;
            //     }
            // }
        }
    }
}
