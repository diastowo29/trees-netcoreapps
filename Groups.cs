using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Authenticators;
using System.Net.Mail;

namespace newConsole {
	class Groups {
		public Dictionary<string,string> searchGroup (string groupMm, string groupDh, string groupAh, string zendeskUsername, string zendeskPassword) {
            Dictionary<string,string> groupList = new Dictionary<string,string>();
            CallingApi myCall = new CallingApi();
            string srcGroup = "/api/v2/search.json?query=type:group";
            StringBuilder srcParameter = new StringBuilder();
            srcParameter.Append(srcGroup);

            srcParameter.Append(" name:\"").Append(groupMm).Append("\"");
            srcParameter.Append(" name:\"").Append(groupDh).Append("\"");
            srcParameter.Append(" name:\"").Append(groupAh).Append("\"");

            JObject searchResponse = JObject.Parse(myCall.callApi(srcParameter.ToString(), zendeskUsername, zendeskPassword));
            JArray arrayResult = (JArray)searchResponse["results"];
            bool groupMMFound = false;
            bool groupDHFound = false;
            bool groupAHFound = false;
            try {
                for (int i=0; i<arrayResult.Count; i++) {
                    if (arrayResult[i]["name"].ToString().ToLower() == groupMm.ToLower()) {
                        // Console.WriteLine("groupMMid found");
                        groupList.Add("groupMMid", arrayResult[i]["id"].ToString());
                        groupMMFound = true;
                    }
                    if (arrayResult[i]["name"].ToString().ToLower() == groupDh.ToLower()) {
                        // Console.WriteLine("groupDHid found");
                        groupList.Add("groupDHid", arrayResult[i]["id"].ToString());
                        groupDHFound = true;
                    }
                    if (arrayResult[i]["name"].ToString().ToLower() == groupAh.ToLower()) {
                        // Console.WriteLine("groupAHid found");
                        groupList.Add("groupAHid", arrayResult[i]["id"].ToString());
                        groupAHFound = true;
                    }
                }
            } catch {
                Console.WriteLine("there is error on group..");
            }

            if (!groupMMFound) {
                // Console.WriteLine("groupMMid not found");
                groupList.Add("groupMMid", "0");
            }

            if (!groupDHFound) {
                // Console.WriteLine("groupDHid not found");
                groupList.Add("groupDHid", "0");
            }

            if (!groupAHFound) {
                // Console.WriteLine("groupAHid not found");
                groupList.Add("groupAHid", "0");
            }
            
            return groupList;
        }

        public string createGroup (string groupName, string zendeskUsername, string zendeskPassword) {
        	string groupId = "";
        	CallingApi myCall = new CallingApi();
        	string createGroupApi = "/api/v2/groups.json";
        	Dictionary<string,string> groupProp = new Dictionary<string,string>();
        	groupProp.Add("name", groupName);
        	Dictionary<string,object> group = new Dictionary<string,object>();
        	group.Add("group", groupProp);
        	JObject createGroupResponse = JObject.Parse(myCall.callApiPost(JsonConvert.SerializeObject(group), createGroupApi, zendeskUsername, zendeskPassword));
        	groupId = createGroupResponse["group"]["id"].ToString();

        	return groupId;
        }

        public void checkGroupMemberships (string userId, string groupId, List<Dictionary<string,string>> agentGroupList, string zendeskUsername, string zendeskPassword) {
            CallingApi myCall = new CallingApi();
            string showMembershipApi = "/api/v2/groups/" + groupId + "/memberships.json";
            JObject membershipList = JObject.Parse(myCall.callApi(showMembershipApi, zendeskUsername, zendeskPassword));
            JArray memberArray = (JArray)membershipList["group_memberships"];
            StringBuilder deleteList = new StringBuilder();
            bool isEmpty = true;
            if (memberArray.Count > 0) {
                for (int i=0; i<memberArray.Count; i++) {
                    if (memberArray[i]["user_id"].ToString() != userId) {
                        isEmpty = false;
                        deleteList.Append(memberArray[i]["id"].ToString()).Append(",");
                    }
                }
            }
            if (!isEmpty) {
                deleteManyMembers(deleteList.ToString(), zendeskUsername, zendeskPassword);
            }
        }

        public void addMembership (List<Dictionary<string,string>> memberList, List<Dictionary<string,string>> agentGroupList, string zendeskUsername, string zendeskPassword) {
            CallingApi myCall = new CallingApi();
            string createManyMembershipApi = "/api/v2/group_memberships/create_many.json";
            if (memberList.Count > 0) {
                Dictionary<string,object> groupMembership = new Dictionary<string,object>();
                groupMembership.Add("group_memberships", memberList);
                myCall.callApiPost(JsonConvert.SerializeObject(groupMembership), createManyMembershipApi, zendeskUsername, zendeskPassword);
            }
            for (int i=0; i<memberList.Count; i++) {
                checkGroupMemberships(memberList[i]["user_id"].ToString(), memberList[i]["group_id"].ToString(), agentGroupList, zendeskUsername, zendeskPassword);
            }
        }

        public void deleteManyMembers (string deleteParameter, string zendeskUsername, string zendeskPassword) {
            CallingApi myCall = new CallingApi();
            string deleteMembersApi = "/api/v2/group_memberships/destroy_many.json?ids=" + deleteParameter;
            myCall.callApiDelete(deleteMembersApi, zendeskUsername, zendeskPassword);
        }

        public void deleteAllMember (List<string> groupIds, string zendeskUsername, string zendeskPassword) {
            Console.WriteLine("===== DELETE ALL MEMBER =====");
            CallingApi myCall = new CallingApi();
            string getMemberApi = "";
            string deleteManyMemberApi = "/api/v2/group_memberships/destroy_many.json?ids=";
            StringBuilder deleteParameter = new StringBuilder();
            deleteParameter.Append(deleteManyMemberApi);
            bool isEmpty = true;

            for (int i=0; i<groupIds.Count; i++) {
                if (groupIds[i] != "0") {
                    getMemberApi = "/api/v2/groups/" + groupIds[i] + "/memberships.json";
                    JObject allMember = JObject.Parse(myCall.callApi(getMemberApi, zendeskUsername, zendeskPassword));
                    JArray allMemberArray = (JArray)allMember["group_memberships"];
                    if (allMemberArray.Count > 0) {
                        for (int j=0; j<allMemberArray.Count; j++) {
                            deleteParameter.Append(allMemberArray[j]["id"].ToString()).Append(",");
                            isEmpty = false;
                        }
                    }
                }
            }
            if (!isEmpty) {
                myCall.callApiDelete(deleteParameter.ToString(), zendeskUsername, zendeskPassword);
            }
        }

        public void checkAgentMembership (string userId, string groupId, List<Dictionary<string, string>> entries, string type, string zendeskUsername, string zendeskPassword) {
            CallingApi myCall = new CallingApi();
            if (userId != "0") {
                string getAgentMemberApi = "/api/v2/users/" + userId + "/group_memberships.json?include=groups,users";
                StringBuilder deleteParameter = new StringBuilder();
                JObject memberList = JObject.Parse(myCall.callApi(getAgentMemberApi, zendeskUsername, zendeskPassword));
                JArray memberlistArray = (JArray)memberList["group_memberships"];
                JArray groupsArray = (JArray)memberList["groups"];
                JArray usersArray = (JArray)memberList["users"];
                if (groupsArray.Count > 1) {
                    for(int i=0; i<groupsArray.Count; i++) {
                        bool groupMMFound = false;
                        string falseGroupId = "0";
                        bool groupDHFound = false;
                        bool groupAHFound = false;
                        if (groupsArray[i]["name"].ToString().Contains("Group D ")) {
                            // Console.WriteLine("GROUP D");
                            /*MM*/
                            if (groupsArray[i]["name"].ToString().Contains("MM")) {
                                // Console.WriteLine("Group MM");
                                for (int j=0; j<entries.Count; j++) {
                                    if (entries[j]["MM"].ToString() == usersArray[0]["name"].ToString()) {
                                        if (groupsArray[i]["name"].ToString().Contains(entries[j]["Branch Code"].ToString())) {
                                            groupMMFound = true;
                                        }
                                    }
                                }
                                if (!groupMMFound) {
                                    falseGroupId = groupsArray[i]["id"].ToString();
                                    for (int j=0; j<memberlistArray.Count; j++) {
                                        if (memberlistArray[j]["group_id"].ToString() == falseGroupId) {
                                            deleteParameter.Append(memberlistArray[i]["id"].ToString()).Append(",");
                                            falseGroupId = "0";
                                        }
                                    }
                                }
                            }

                            /*DH*/
                            if (groupsArray[i]["name"].ToString().Contains("Dept Head")) {
                                // Console.WriteLine("Group DH");
                                for (int j=0; j<entries.Count; j++) {
                                    // Console.WriteLine(JsonConvert.SerializeObject(entries[j]));
                                    if (entries[j]["Dept. Head"].ToString() == usersArray[0]["name"].ToString()) {
                                        if (groupsArray[i]["name"].ToString().Contains(entries[j]["Region"].ToString())) {
                                            groupDHFound = true;
                                        }
                                    }
                                }
                                if (!groupDHFound) {
                                    falseGroupId = groupsArray[i]["id"].ToString();
                                    for (int j=0; j<memberlistArray.Count; j++) {
                                        if (memberlistArray[j]["group_id"].ToString() == falseGroupId) {
                                            deleteParameter.Append(memberlistArray[i]["id"].ToString()).Append(",");
                                            falseGroupId = "0";
                                        }
                                    }
                                }
                            }

                            /*AH*/
                            if (groupsArray[i]["name"].ToString().Contains("Area Head")) {
                                // Console.WriteLine("Group AH");
                                for (int j=0; j<entries.Count; j++) {
                                    if (entries[j]["NAMA Area Head"].ToString() == usersArray[0]["name"].ToString()) {
                                        if (groupsArray[i]["name"].ToString().Contains(entries[j]["AREA"].ToString())) {
                                            groupAHFound = true;
                                        }
                                    }
                                }
                                if (!groupAHFound) {
                                    falseGroupId = groupsArray[i]["id"].ToString();
                                    for (int j=0; j<memberlistArray.Count; j++) {
                                        if (memberlistArray[j]["group_id"].ToString() == falseGroupId) {
                                            deleteParameter.Append(memberlistArray[i]["id"].ToString()).Append(",");
                                            falseGroupId = "0";
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (deleteParameter.Length > 0) {
                        deleteManyMembers(deleteParameter.ToString(), zendeskUsername, zendeskPassword);
                    }
                }
            }
        }
	}
}