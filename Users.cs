using System;
using System.Net.Mail;
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

namespace newConsole {
	class Users {
		public List<JToken> checkDealer(List<JToken> userCreated, List<Dictionary<string, string>> entries,string orgId, bool doneOwner, bool donePicOwner, bool donePicOutlet, Dictionary<string,string> currEntries, string zendeskUsername, string zendeskPassword) {
            string kuadran = currEntries["Kuadran"].ToString();
            string jenisGroup = currEntries["Jenis Group"].ToString();
            string branchCode = currEntries["Branch Code"].ToString();

            string dealerName = currEntries["Dealer Name"].ToString();
            string dealerId = currEntries["Dealer ID"].ToString();
            string emailPicOwner = currEntries["Alamat email PIC Owner"].ToString();
            string emailOwner = currEntries["Alamat email owner"].ToString();
            string emailPicOutlet = currEntries["Alamat email PIC Outlet"].ToString();
            List<JToken> userJustCreated = new List<JToken>();

            Console.WriteLine("doneOwner: {0}, donePicOwner: {1}, donePicOutlet: {2}", doneOwner, donePicOwner, donePicOutlet);
			CallingApi callingApi = new CallingApi();
			string srcUser = "/api/v2/search.json?query=type:user";
			StringBuilder srcParameter = new StringBuilder();
			srcParameter.Append(srcUser);

            List<Dictionary<string,string>> orgMember = new List<Dictionary<string,string>>();
            Dictionary<string,string> members = new Dictionary<string,string>();

            string createUserApi = "/api/v2/users.json";

            if (emailOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailOwner).Append("\"");
            }
            if (emailPicOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOwner).Append("\"");
            }
            if (emailPicOutlet != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOutlet).Append("\"");
            }
            
            string srcUserResponse = callingApi.callApi(srcParameter.ToString(), zendeskUsername, zendeskPassword);
            JObject srcUserJoResponse = JObject.Parse(srcUserResponse);
            JArray srcUserArray = (JArray)srcUserJoResponse["results"];
            string ownerFound = "0";
            string picOwnerFound = "0";
            string picOutletFound = "0";

            JObject ownerDict = new JObject();
            JObject picOwnerDict = new JObject();
            JObject picOutletDict = new JObject();

            string userId = "";

            for (int i=0; i<srcUserArray.Count; i++) {
        		if (srcUserArray[i]["email"].ToString().ToLower() == emailOwner.ToLower()) {
        			ownerFound = srcUserArray[i]["id"].ToString();
                    ownerDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
        		}

        		if (srcUserArray[i]["email"].ToString().ToLower() == emailPicOwner.ToLower()) {
        			picOwnerFound = srcUserArray[i]["id"].ToString();
                    picOwnerDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
        		}

        		if (srcUserArray[i]["email"].ToString().ToLower() == emailPicOutlet.ToLower()) {
        			picOutletFound = srcUserArray[i]["id"].ToString();
                    picOutletDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
        		}
            }

            if (!emailOwner.Equals("#N/A")) {
        		Dictionary<string,string> ownerUserCustomField = new Dictionary<string,string>();
                int atPos = emailOwner.IndexOf("@");
                string emailUsername = emailOwner.Substring(0, atPos);
        		Dictionary<string,object> ownerUserField = new Dictionary<string,object>();
                ownerUserField.Add("name", emailUsername);
                ownerUserField.Add("email", emailOwner);
                // ownerUserField.Add("organization_id", orgId);
                ownerUserCustomField.Add("jabatan", "owner");
                ownerUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> ownerUser = new Dictionary<string,object>();
                if (doneOwner) {
                    try {
                        if (!ownerDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            ownerUserCustomField.Add("dealer_id", ownerDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!ownerDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            ownerUserCustomField.Add("branch_id", ownerDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!ownerDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            ownerUserCustomField.Add("kuadran", ownerDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!ownerDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            ownerUserCustomField.Add("dealer_name", ownerDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    } catch (NullReferenceException e) {
                        Console.WriteLine(e);
                        Console.WriteLine("===== USER CREATED BUT CANNOT BE FOUND =====");
                        for (int i=0; i<userCreated.Count; i++) {
                            if (userCreated[i]["email"].ToString().ToLower() == emailOwner.ToLower()) {
                                ownerFound = userCreated[i]["id"].ToString();
                                if (!userCreated[i]["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                                    ownerUserCustomField.Add("dealer_id", userCreated[i]["user_fields"]["dealer_id"] +  ";" + dealerId);
                                }
                                if (!userCreated[i]["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                                    ownerUserCustomField.Add("branch_id", userCreated[i]["user_fields"]["branch_id"] +  ";" + branchCode);
                                }
                                if (!userCreated[i]["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                                    ownerUserCustomField.Add("kuadran", userCreated[i]["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                                }
                                if (!userCreated[i]["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                                    ownerUserCustomField.Add("dealer_name", userCreated[i]["user_fields"]["dealer_name"] +  ";" + dealerName);
                                }
                            }
                        }
                    }
                    ownerUserField.Add("user_fields", ownerUserCustomField);
                    ownerUser.Add("user", ownerUserField);

                    string updateUserApi = "/api/v2/users/" + ownerFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(ownerUser), updateUserApi, zendeskUsername, zendeskPassword);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();

                    List<string> tagsList = new List<string>();
                    string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                    tagsList.Add(tagsInput);
                    Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                    tags.Add("tags", tagsList);
                    string addTagsApi = "/api/v2/users/" + ownerFound + "/tags.json";
                    string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi, zendeskUsername, zendeskPassword);

                } else {
                    if (ownerFound == "0") {
                        ownerUserCustomField.Add("dealer_id", dealerId);
                        ownerUserCustomField.Add("branch_id", branchCode);
                        ownerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        ownerUserCustomField.Add("dealer_name", dealerName);
                        ownerUserField.Add("user_fields", ownerUserCustomField);
                        ownerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                        ownerUser.Add("user", ownerUserField);
                        string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(ownerUser), createUserApi, zendeskUsername, zendeskPassword);
                        JObject createdUserResponse = JObject.Parse(createUser);
                        userJustCreated.Add(createdUserResponse["user"]);
                        userId = createdUserResponse["user"]["id"].ToString();
                        if (emailOwner == emailPicOwner) {
                            donePicOwner = true;
                            picOwnerFound = userId;
                            picOwnerDict = (JObject)createdUserResponse["user"];
                        } else if (emailOwner == emailPicOutlet) {
                            donePicOutlet = true;
                            picOutletFound = userId;
                            picOutletDict = (JObject)createdUserResponse["user"];
                        }
                    } else {
                        ownerUserCustomField.Add("dealer_id", dealerId);
                        ownerUserCustomField.Add("branch_id", branchCode);
                        ownerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        ownerUserCustomField.Add("dealer_name", dealerName);
                        ownerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);

                        ownerUserField.Add("user_fields", ownerUserCustomField);
                        ownerUser.Add("user", ownerUserField);

                        string updateUserApi = "/api/v2/users/" + ownerFound + ".json";
                        string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(ownerUser), updateUserApi, zendeskUsername, zendeskPassword);
                        JObject updateUserResponse = JObject.Parse(updateUser);
                        userId = updateUserResponse["user"]["id"].ToString();
                    }
                }
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }

            if (!emailPicOwner.Equals("#N/A")) {
        		Dictionary<string,string> picOwnerUserCustomField = new Dictionary<string,string>();
                int atPos = emailPicOwner.IndexOf("@");
                string emailUsername = emailPicOwner.Substring(0, atPos);
                Dictionary<string,object> picOwnerUserField = new Dictionary<string,object>();
                picOwnerUserField.Add("name", emailUsername);
                picOwnerUserField.Add("email", emailPicOwner);
                // picOwnerUserField.Add("organization_id", orgId);
                picOwnerUserCustomField.Add("jabatan", "pic_owner");
                picOwnerUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> picOwnerUser = new Dictionary<string,object>();
                if (donePicOwner) {
                    try {
                        if (!picOwnerDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            picOwnerUserCustomField.Add("dealer_id", picOwnerDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!picOwnerDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            picOwnerUserCustomField.Add("branch_id", picOwnerDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!picOwnerDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            picOwnerUserCustomField.Add("kuadran", picOwnerDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!picOwnerDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            picOwnerUserCustomField.Add("dealer_name", picOwnerDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    } catch (NullReferenceException e) {
                        Console.WriteLine(e);
                        Console.WriteLine("===== USER CREATED BUT CANNOT BE FOUND =====");
                        for (int i=0; i<userCreated.Count; i++) {
                            if (userCreated[i]["email"].ToString().ToLower() == emailPicOwner.ToLower()) {
                                picOwnerFound = userCreated[i]["id"].ToString();
                                if (!userCreated[i]["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                                    picOwnerUserCustomField.Add("dealer_id", userCreated[i]["user_fields"]["dealer_id"] +  ";" + dealerId);
                                }
                                if (!userCreated[i]["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                                    picOwnerUserCustomField.Add("branch_id", userCreated[i]["user_fields"]["branch_id"] +  ";" + branchCode);
                                }
                                if (!userCreated[i]["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                                    picOwnerUserCustomField.Add("kuadran", userCreated[i]["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                                }
                                if (!userCreated[i]["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                                    picOwnerUserCustomField.Add("dealer_name", userCreated[i]["user_fields"]["dealer_name"] +  ";" + dealerName);
                                }
                            }
                        }
                    }
                    picOwnerUserField.Add("user_fields", picOwnerUserCustomField);
                    picOwnerUser.Add("user", picOwnerUserField);

                    string updateUserApi = "/api/v2/users/" + picOwnerFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOwnerUser), updateUserApi, zendeskUsername, zendeskPassword);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();

                    List<string> tagsList = new List<string>();
                    string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                    tagsList.Add(tagsInput);
                    Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                    tags.Add("tags", tagsList);
                    string addTagsApi = "/api/v2/users/" + picOwnerFound + "/tags.json";
                    string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi, zendeskUsername, zendeskPassword);

                } else {
                    if (picOwnerFound == "0") {
                        picOwnerUserCustomField.Add("dealer_id", dealerId);
                        picOwnerUserCustomField.Add("branch_id", branchCode);
                        picOwnerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOwnerUserCustomField.Add("dealer_name", dealerName);
                        picOwnerUserField.Add("user_fields", picOwnerUserCustomField);
                        picOwnerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                        picOwnerUser.Add("user", picOwnerUserField);
                        string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(picOwnerUser), createUserApi, zendeskUsername, zendeskPassword);
                        JObject createdUserResponse = JObject.Parse(createUser);
                        userJustCreated.Add(createdUserResponse["user"]);
                        userId = createdUserResponse["user"]["id"].ToString();
                        if (emailPicOwner == emailOwner) {
                            doneOwner = true;
                            ownerFound = userId;
                            ownerDict = (JObject)createdUserResponse["user"];
                        } else if (emailPicOwner == emailPicOutlet) {
                            donePicOutlet = true;
                            picOutletFound = userId;
                            picOutletDict = (JObject)createdUserResponse["user"];
                        }
                    } else {
                        picOwnerUserCustomField.Add("dealer_id", dealerId);
                        picOwnerUserCustomField.Add("branch_id", branchCode);
                        picOwnerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOwnerUserCustomField.Add("dealer_name", dealerName);
                        picOwnerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);

                        picOwnerUserField.Add("user_fields", picOwnerUserCustomField);
                        picOwnerUser.Add("user", picOwnerUserField);

                        string updateUserApi = "/api/v2/users/" + picOwnerFound + ".json";
                        string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOwnerUser), updateUserApi, zendeskUsername, zendeskPassword);
                        JObject updateUserResponse = JObject.Parse(updateUser);
                        userId = updateUserResponse["user"]["id"].ToString();
                    }
                }
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }

            if (!emailPicOutlet.Equals("#N/A")) {
        		Dictionary<string,string> picOutletUserCustomField = new Dictionary<string,string>();
                int atPos = emailPicOutlet.IndexOf("@");
                string emailUsername = emailPicOutlet.Substring(0, atPos);
                Dictionary<string,object> picOutletUserField = new Dictionary<string,object>();
                picOutletUserField.Add("name", emailUsername);
                picOutletUserField.Add("email", emailPicOutlet);
                // picOutletUserField.Add("organization_id", orgId);
                picOutletUserCustomField.Add("jabatan", "pic_outlet");
                picOutletUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> picOutletUser = new Dictionary<string,object>();
                if (donePicOutlet) {
                    try {
                        if (!picOutletDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            picOutletUserCustomField.Add("dealer_id", picOutletDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!picOutletDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            picOutletUserCustomField.Add("branch_id", picOutletDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!picOutletDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            picOutletUserCustomField.Add("kuadran", picOutletDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!picOutletDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            picOutletUserCustomField.Add("dealer_name", picOutletDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    } catch (NullReferenceException e) {
                        Console.WriteLine(e);
                        Console.WriteLine("===== USER CREATED BUT CANNOT BE FOUND =====");
                        for (int i=0; i<userCreated.Count; i++) {
                            if (userCreated[i]["email"].ToString().ToLower() == emailPicOutlet.ToLower()) {
                                picOutletFound = userCreated[i]["id"].ToString();
                                if (!userCreated[i]["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                                    picOutletUserCustomField.Add("dealer_id", userCreated[i]["user_fields"]["dealer_id"] +  ";" + dealerId);
                                }
                                if (!userCreated[i]["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                                    picOutletUserCustomField.Add("branch_id", userCreated[i]["user_fields"]["branch_id"] +  ";" + branchCode);
                                }
                                if (!userCreated[i]["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                                    picOutletUserCustomField.Add("kuadran", userCreated[i]["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                                }
                                if (!userCreated[i]["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                                    picOutletUserCustomField.Add("dealer_name", userCreated[i]["user_fields"]["dealer_name"] +  ";" + dealerName);
                                }
                            }
                        }
                    }
                    picOutletUserField.Add("user_fields", picOutletUserCustomField);
                    picOutletUser.Add("user", picOutletUserField);

                    string updateUserApi = "/api/v2/users/" + picOutletFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOutletUser), updateUserApi, zendeskUsername, zendeskPassword);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();

                    List<string> tagsList = new List<string>();
                    string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                    tagsList.Add(tagsInput);
                    Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                    tags.Add("tags", tagsList);
                    string addTagsApi = "/api/v2/users/" + picOutletFound + "/tags.json";
                    string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi, zendeskUsername, zendeskPassword);

                } else {
                    if (picOutletFound == "0") {
                        picOutletUserCustomField.Add("dealer_id", dealerId);
                        picOutletUserCustomField.Add("branch_id", branchCode);
                        picOutletUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOutletUserCustomField.Add("dealer_name", dealerName);
                        picOutletUserField.Add("user_fields", picOutletUserCustomField);
                        picOutletUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                        picOutletUser.Add("user", picOutletUserField);
                        string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(picOutletUser), createUserApi, zendeskUsername, zendeskPassword);
                        JObject createdUserResponse = JObject.Parse(createUser);
                        userJustCreated.Add(createdUserResponse["user"]);
                        userId = createdUserResponse["user"]["id"].ToString();
                        if (emailPicOutlet == emailOwner) {
                            doneOwner = true;
                            ownerFound = userId;
                            ownerDict = (JObject)createdUserResponse["user"];
                        } else if (emailPicOutlet == emailPicOwner) {
                            donePicOwner = true;
                            picOwnerFound = userId;
                            picOwnerDict = (JObject)createdUserResponse["user"];
                        }
                    } else {
                        picOutletUserCustomField.Add("dealer_id", dealerId);
                        picOutletUserCustomField.Add("branch_id", branchCode);
                        picOutletUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOutletUserCustomField.Add("dealer_name", dealerName);
                        picOutletUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);

                        picOutletUserField.Add("user_fields", picOutletUserCustomField);
                        picOutletUser.Add("user", picOutletUserField);

                        string updateUserApi = "/api/v2/users/" + picOutletFound + ".json";
                        string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOutletUser), updateUserApi, zendeskUsername, zendeskPassword);
                        JObject updateUserResponse = JObject.Parse(updateUser);
                        userId = updateUserResponse["user"]["id"].ToString();
                    }
                }
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }
            if (orgMember.Count > 0) {
                createOrgMemberships(orgMember, zendeskUsername, zendeskPassword);
                for (int i=0; i<orgMember.Count; i++) {
                    checkUserMemberships(orgMember[i], entries, zendeskUsername, zendeskPassword);
                }
            }
            return userJustCreated;
        }
        public void createOrgMemberships(List<Dictionary<string,string>> orgMember, string zendeskUsername, string zendeskPassword) {
            CallingApi callingApi = new CallingApi();
            string createManyOrgMembershipApi = "/api/v2/organization_memberships/create_many.json";
            Dictionary<string,object> orgMembership = new Dictionary<string,object>();
            orgMembership.Add("organization_memberships", orgMember);
            string createMembership = callingApi.callApiPost(JsonConvert.SerializeObject(orgMembership), createManyOrgMembershipApi, zendeskUsername, zendeskPassword);
        }

        public void checkUserMemberships(Dictionary<string,string> userMember, List<Dictionary<string, string>> entries, string zendeskUsername, string zendeskPassword) {
            CallingApi callingApi = new CallingApi();
            string checkMembershipApi = "/api/v2/users/" + userMember["user_id"].ToString() + "/organization_memberships.json?include=organizations,users";
            string checkMembership = callingApi.callApi(checkMembershipApi, zendeskUsername, zendeskPassword);
            JObject checkMembershipResponse = JObject.Parse(checkMembership);

            JArray membershipArray = (JArray)checkMembershipResponse["organization_memberships"];
            JArray organizationArray = (JArray)checkMembershipResponse["organizations"];
            JArray usersArray = (JArray)checkMembershipResponse["users"];

            StringBuilder deleteList = new StringBuilder();

            if (membershipArray.Count > 1) {
                string userEmail = usersArray[0]["email"].ToString();
                Console.WriteLine("membershipArray count more than 1");
                for (int i=0; i<membershipArray.Count; i++) {
                    bool orgFound = false;
                    Console.WriteLine(organizationArray[i]["name"].ToString());
                    if (organizationArray[i]["name"].ToString().StartsWith("D ")) {
                        string orgName = organizationArray[i]["name"].ToString();
                        string[] orgBranchCode = orgName.Split(' ');

                        for (int j=0; j<entries.Count; j++) {
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email PIC Owner"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email owner"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email PIC Outlet"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                        }
                    }
                    if (!orgFound) {
                        deleteList.Append(membershipArray[i]["id"].ToString()).Append(",");
                    }
                }
                // Console.WriteLine(JsonConvert.SerializeObject(deleteList));
                if (deleteList.ToString().Length > 0) {
                    deleteMembership(deleteList.ToString().Remove((deleteList.ToString().Length)-1, 1), zendeskUsername, zendeskPassword);
                }
            }
        }

        public void deleteMembership(string deleteList, string zendeskUsername, string zendeskPassword) {
            string deleteManyOrgMemberApi = "/api/v2/organization_memberships/destroy_many.json?ids=" + deleteList;
            CallingApi callingApi = new CallingApi();
            callingApi.callApiDelete(deleteManyOrgMemberApi, zendeskUsername, zendeskPassword);
        }

        public string createAgent(string userName, string userEmail, string groupId, string customRoleId, string zendeskUsername, string zendeskPassword) {
            string userid = "";
            string createUserApi = "/api/v2/users.json";
            CallingApi myCall = new CallingApi();
            Dictionary<string,string> agentProp = new Dictionary<string,string>();
            agentProp.Add("name", userName);
            agentProp.Add("email", userEmail);
            agentProp.Add("role", "agent");
            agentProp.Add("custom_role_id", customRoleId);
            agentProp.Add("default_group_id", groupId);
            Dictionary<string,object> agent = new Dictionary<string,object>();
            agent.Add("user", agentProp);
            JObject createAgentResponse = JObject.Parse(myCall.callApiPost(JsonConvert.SerializeObject(agent), createUserApi, zendeskUsername, zendeskPassword));
            try {
                userid = createAgentResponse["user"]["id"].ToString();
            } catch {
                Console.WriteLine("error when create agent..");
                userid = "0";
            }

            return userid;
        }

        public Dictionary<string,string> searchAgent (string emailMm, string emaiLDh, string emailAh, string zendeskUsername, string zendeskPassword) {
            Dictionary<string,string> agentList = new Dictionary<string,string>();
            CallingApi myCall = new CallingApi();
            string srcUser = "/api/v2/search.json?query=type:user";
            StringBuilder srcParameter = new StringBuilder();
            srcParameter.Append(srcUser);

            if (!emailMm.Contains("#N/A")) {
                srcParameter.Append(" email:\"").Append(emailMm).Append("\"");
            }
            if (!emaiLDh.Contains("#N/A")) {
                srcParameter.Append(" email:\"").Append(emaiLDh).Append("\"");
            }
            if (!emailAh.Contains("#N/A")) {
                srcParameter.Append(" email:\"").Append(emailAh).Append("\"");
            }

            bool userMmFound = false;
            bool userDhFound = false;
            bool userAhFound = false;

            JObject searchResponse = JObject.Parse(myCall.callApi(srcParameter.ToString(), zendeskUsername, zendeskPassword));
            JArray arrayResult = (JArray)searchResponse["results"];
            for (int i=0; i<arrayResult.Count; i++) {
                if (arrayResult[i]["email"].ToString().ToLower() == emailMm.ToLower()) {
                    agentList.Add("userMmId", arrayResult[i]["id"].ToString());
                    userMmFound = true;
                }
                if (arrayResult[i]["email"].ToString().ToLower() == emaiLDh.ToLower()) {
                    agentList.Add("userDhId", arrayResult[i]["id"].ToString());
                    userDhFound = true;
                }
                if (arrayResult[i]["email"].ToString().ToLower() == emailAh.ToLower()) {
                    agentList.Add("userAhId", arrayResult[i]["id"].ToString());
                    userAhFound = true;
                }
            }

            if (!userMmFound) {
                agentList.Add("userMmId", "0");
            }
            if (!userDhFound) {
                agentList.Add("userDhId", "0");
            }
            if (!userAhFound) {
                agentList.Add("userAhId", "0");
            }

            return agentList;
        }
	}
}