import requests as r
import xlsxwriter as xls

# Aux method library
def create_xlsx():
    resultset = xls.Workbook("users_result_set.xlsx")
    return resultset

def user_inputs():
    #lists
    salesOrgls = ["A001", "B001", "C001", "D001", "E001", "F001", "L001"]
    applicationls = ["CADI", "CSPLUS", "SKILLABLE"]
    rolels = [["ADMIN","APP_OFT_SUP_USER","APP_OFT_USER","APP_ONT_SUP_USER","APP_ONT_USER","APP_ORDER_USER","APP_USER","BO_OFT_ADMIN","BO_OFT_USER","BO_ONT_ADMIN","BO_ONT_USER","BO_USER"],
    ["ADMIN","BO_USER","LOYALTY_MGR","MANAGER","OWNER"],["ADMIN","CHAIN_OWNER","CONTENT_MNG","MANAGER","REPRESENTATIVE","STAFF"]] #roles for each app. the distinction between apps is based on the position of the app in the list "applicationls"
    #inputs for user
    port = input("please insert the port associated with the cx-iam-services.live.services:80 (inside the \"config\" file in the ssh folder)")
    applicationName = input("insert the desired application (select only 1):\n" + ",\n".join(applicationls) + "\n")
    roles = input(
        "insert which roles you want to retrieve (to insert multiple, please add a whitespace between each value, or type \"all\" to select all):\n" + ",\n".join(rolels[applicationls.index(applicationName.upper())]) + "\n")
    salesOrg = input(
        "input the desired salesOrgs (to insert multiple, please add a whitespace between each value):\n" + ",\n".join(salesOrgls) + "\n")
    #converting user string input to list
    sales_org_result = salesOrg.upper().split()
    roles_result = roles.upper().split()
    return port, applicationName, sales_org_result,roles_result

def write_roles(currentSheet,roles, email, row,roles_result,salesOrg):
    for k in range(len(roles)):
                # this if is necessary because the request sends some results that dont match the configurations made in the url
                if (
                roles[k]["role"]["applicationCode"] == applicationName.upper() and "salesOrgCode" not in roles[k]) or (
                roles[k]["role"]["applicationCode"] == applicationName.upper() and roles[k]["salesOrgCode"] == salesOrg.upper()): 
                    if(str(roles_result[0]) == "ALL" and len(roles_result) == 1):
                        currentSheet.write(row, 0, email)
                        currentSheet.write(row, 1,roles[k]["role"]["roleName"])
                        row += 1
                    else:
                        if(str(roles[k]["role"]["roleName"]).upper() in roles_result):
                            currentSheet.write(row, 0, email)
                            currentSheet.write(row, 1,roles[k]["role"]["roleName"])
                            row += 1
    # print("loop out \"write roles\" \n")
    return currentSheet, email, row


def user_roles_to_excel(port,applicationName, salesOrg,resultset,roles_result):
    row = 0
    currentSheet = resultset.add_worksheet(salesOrg)
    currentSheet.write(row,0,"email")
    currentSheet.write(row,1,"role")
    row += 1
    url = 'http://localhost:{port}/services/salesorgs/{salesOrg}/users?type=INTERNAL&applicationCode={application}&page={page}'
    numPages = r.get(url.format(
        port = port,salesOrg = salesOrg, application = applicationName, page = 0 )).json()["totalPages"] # number of pages in the response

    # filtering data from JSON response
    for i in range(numPages): # used to iterate between the response pages
        page = r.get(
            url.format(port = port,salesOrg = salesOrg, application = applicationName, page = i )).json()
        users = page["content"]
        for j in range(len(users)): # used to iterate between the users in the response
            email = users[j]["emails"][0]["address"]
            if "userSalesOrgRoles" in users[j]["roles"] and len(users[j]["roles"]["userSalesOrgRoles"]) > 0: currentSheet, email, row = write_roles(currentSheet, users[j]["roles"]["userSalesOrgRoles"], email, row, roles_result,salesOrg)
            if "userRoles" in users[j]["roles"] and len(users[j]["roles"]["userRoles"]) > 0: currentSheet, email, row = write_roles(currentSheet, users[j]["roles"]["userRoles"], email, row, roles_result,salesOrg)
    #     print("loop out \"roles_to_excell inner loop\" \n")    
    # print("loop out \"roles_to_excell outer loop\" \n")
def multiple_sales_org(port, applicationName,salesOrgls,resultset,roles_result):
    for i in range(len(salesOrgls)):user_roles_to_excel(port,applicationName, salesOrgls[i],resultset,roles_result)

#End of Method Library






# MAIN
 
# Creating Excel file and sheets
resultset = create_xlsx()

# Getting user inputs
port, applicationName, salesOrgls,roles_result = user_inputs()
print("\nExtracting roles to \"users_result_set.xlsx\", please wait...")
# case of result with multiple sales orgs
if len(salesOrgls) > 1: multiple_sales_org(port, applicationName,salesOrgls, resultset,roles_result)  

# case of a single sales org
else: user_roles_to_excel(port, applicationName,salesOrgls[0], resultset,roles_result)
input("\nOperation complete! press \"enter\" to close this screen") 
resultset.close()
