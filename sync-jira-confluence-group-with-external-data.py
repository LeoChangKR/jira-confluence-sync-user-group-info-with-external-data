from google.oauth2 import service_account
from atlassian import Confluence
from atlassian import Jira
from pyppeteer import launch
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
import base64
import requests
import json
import time
import codecs
import asyncio
from io import StringIO # Using StringBuilder from https://appdividend.com/2022/01/21/string-builder-equivalent-in-python/
 
class StringBuilder:
    _file_str = None
    def __init__(self):
        self._file_str = StringIO()
    def Add(self, str):
        self._file_str.write(str)
    def __str__(self):
        return self._file_str.getvalue()
 
string_builder = StringBuilder()
 
# Global variables
wiki_group_and_member_data = []
jira_group_and_member_data = []
sorted_categories = []
wiki_script_url = "$YOUR_URL$/admin/power-groovy/script-launcher.action"
#power-groovy is a free plugin. It's a simple script console without any options, just basically runs anything inside that console.
wiki_username = $ADMIN_ID$
wiki_password = $ADMIN_PW$
today = time.strftime('%Y-%m-%d', time.localtime(time.time()))
 
# Google Sheets authentication
credentials = service_account.Credentials.from_service_account_file(
    "$YOUR_CREDENTIALS$",
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
)
client = gspread.Client(credentials)
 
# Open the spreadsheet using the provided URL
spreadsheet_url = "$YOUR_SHEET$"
spreadsheet = client.open_by_url(spreadsheet_url)
 
#Access Jira
jira = Jira(
    url='$YOUR_URL$',
    username= $ADMIN_ID$
    password= $ADMIN_PW$)
 
# Function to clear existing data in the worksheet (except row number 1)
def clear_worksheet(worksheet):
    existing_data = worksheet.get_all_values()
    existing_data[1:] = []
    worksheet.clear()
    worksheet.update(existing_data)
 
# Function to get all groups from Confluence
def get_all_groups_confluence():
    # Confluence Server information
    confluence_url = '$YOUR_URL$'
    personal_access_token = $TOKEN$
 
    # Create Confluence object
    confluence = Confluence(url=confluence_url, token=personal_access_token)
 
    groups = []
    start = 0
    limit = 1000
    while True:
        # Use params in the get method to include the limit and start parameters
        response = confluence.get("rest/api/group", params={'limit': limit, 'start': start})
        groups += response['results']
        if response['size'] < limit:
            break
        start += limit
 
    return groups
 
# Function to get members of a group from Confluence
def get_group_members_confluence(groupname):
    if '/' in groupname:
        print("Groupname should not include a slash(/). Please provide a valid groupname.")
        return None
 
    # Confluence Server information
    confluence_url = '$YOUR_URL$'
    personal_access_token = $TOKEN$
 
    # Create Confluence object
    confluence = Confluence(url=confluence_url, token=personal_access_token)
 
    members = []
    start = 0
    limit = 200
    while True:
        response = confluence.get(f"rest/api/group/{groupname}/member", params={'start': start, 'limit': limit})
        members += response['results']
        if response['size'] < limit:
            break
        start += limit
 
    return members
 
# 1st code: Fetch data from Confluence and append to "Wiki" worksheet
def display_wiki_groups_and_members():
    global wiki_group_and_member_data
    string_builder.Add(str(today) + " Employee sync Service\n")
 
    groups = get_all_groups_confluence()
 
    wiki_group_and_member_data = []  # Initialize an empty list to hold the group and member data
 
    for group in groups:
        groupname = group['name']
 
        # Debugging - Display group name
        members = get_group_members_confluence(groupname)
        if members:
            member_names = [member['username'] for member in members]
            member_names_str = ", ".join(member_names)
        else:
            member_names_str = ""
 
        # Append group and member data to the list
        wiki_group_and_member_data.append([groupname, member_names_str])
 
    # Append all group and member data to the worksheet
    append_data_to_sheet("Wiki", wiki_group_and_member_data)
 
# Function to get all groups from Jira
def get_all_groups_jira():
    # Jira Server information
    jira_url = '$YOUR_URL$'
    api_url = jira_url + '/rest/api/2'
 
    # Jira ID and password
    jira_username = $ADMIN_ID$
    jira_password = $ADMIN_PW$
 
    groups_url = api_url + '/groups/picker'
    params = {'maxResults': 2147483647}
    response = requests.get(groups_url, params=params, auth=(jira_username, jira_password))
 
    if response.status_code == 200:
        return response.json()['groups']
    else:
        print("Failed to retrieve groups.")
        return []
 
# Function to get members of a group from Jira
def get_group_members_jira(group_name):
    # Jira Server information
    jira_url = '$YOUR_URL$'
    api_url = jira_url + '/rest/api/2'
 
    # Jira ID and password
    jira_username = $ADMIN_ID$
    jira_password = $ADMIN_PW$
 
    members_url = api_url + '/group/member'
    params = {'groupname': group_name, 'includeInactiveUsers': 'false', 'maxResults': 50}
    all_members = []
 
    while True:
        response = requests.get(members_url, params=params, auth=(jira_username, jira_password))
 
        if response.status_code == 200:
            members = [member['name'] for member in response.json()['values']]
            all_members.extend(members)
            if response.json()['isLast']:
                break
            else:
                members_url = response.json()['nextPage']
        else:
            print(f"Failed to retrieve members of group: {group_name}.")
            break
 
    return all_members
 
# 2nd code: Fetch data from Jira and append to "Jira" worksheet
def display_jira_groups_and_members():
    global jira_group_and_member_data
 
    groups = get_all_groups_jira()
 
    jira_group_and_member_data = []  # Initialize an empty list to hold the group and member data
 
    for group in groups:
        group_name = group['name']
        group_members = get_group_members_jira(group_name)
 
        # Append group and member data to the list
        jira_group_and_member_data.append([group_name, ", ".join(group_members)])
 
    # Append all group and member data to the worksheet
    append_data_to_sheet("Jira", jira_group_and_member_data)
 
 
# Function to fetch data from the API and append to appropriate worksheets
# I will mask off some parts because it has sensitive data
def fetch_and_append_groupware_data():
    global sorted_categories
 
    # Set up the request payload
    payload = {
        $YOUR PAYLOAD FOR EXTERNAL DATA$
    }
 
    # Send the POST request to the URL
    url = "$YOUR API URL FOR EXTERNAL DATA$"
    response = requests.post(url, data=payload)
 
    # Check if the request was successful
    if response.status_code == 200:
        data = response.json()
        if data["code"] == "0000":
            categories_and_members = []
 
            # Collect unique category_navigation values and member ids
            for employee in data["data"]:
                staff_level = employee.get("staff_level")
                if staff_level == "test":
                    continue
 
                category_navigation = employee.get("category_navigation")
                member_id = employee.get("id")
 
                # Skip accounts that are not actually a user / test accounts
                if member_id and any(keyword in member_id.lower() for keyword in ["test", "bot"]):
                    continue
 
                # Decrypt category_navigation
                decrypted_category = category_navigation.encode().decode('utf-8')
 
                # Skip groups that include "당직"
                if "당직" in decrypted_category or "휴직" in decrypted_category:
                    continue
 
                # Organize the categories and their corresponding members based on their depth levels, ensuring that each category is associated with the appropriate set of members.
                # If a member is part of category A::B::C::D, the member will be added to each category A, A::B, A::B::C, A::B::C::D, and while doing this, if the category doesn't exist, it creates one.
                # This part is needed because the response from groupware API only gives the final category of each user. This part splits the category using depths(::)
                category_parts = decrypted_category.split("::")
                for i in range(1, len(category_parts) + 1):
                    depth_category = "::".join(category_parts[:i])
                    existing_category = next(
                        (
                            (category, members)
                            for category, members in categories_and_members
                            if category == depth_category
                        ),
                        None
                    )
                    if existing_category:
                        existing_category[1].add(member_id)
                    else:
                        categories_and_members.append([depth_category, {member_id}])
 
            # Sort categories based on depth
            sorted_categories = sorted(categories_and_members, key=lambda x: x[0].split("::"))
 
            # Add jira group and wiki group to sorted categories
            for i in range(len(sorted_categories)):
                category, members = sorted_categories[i]
 
                # Add jira group
                #exceptions for YOUR OWN CASE
                jira_group = category.replace("::", "_").replace("$YOUR OWN CASE$", "").replace("$YOUR OWN CASE$", "JIRA_USERS").replace("$YOUR OWN CASE$", "")

                sorted_categories[i].insert(1, jira_group)
 
                # Add wiki group
                #exceptions for YOUR OWN CASE
                wiki_group = category.replace("::", "_").replace("$YOUR OWN CASE$", "").replace("$YOUR OWN CASE$", "JIRA_USERS").replace("$YOUR OWN CASE$", "")
                sorted_categories[i].insert(2, wiki_group)
 
 
            # Separate categories into different groups
            groups = {
                "$YOUR GROUPS$": [],
                "$YOUR GROUPS 2$": [],
                "$YOUR GROUPS 3$": []
            }
 
            for category, jira_group, wiki_group, members in sorted_categories:
                if "$YOUR CASE 1$" in category:
                    groups["YOUR GROUPS"].append((category, jira_group, wiki_group, members))
                elif category.startswith("$YOUR CASE2$"):
                    groups["$YOUR GROUPS 2$"].append((category, jira_group, wiki_group, members))
 
            # Iterate over the groups and append data to the corresponding sheets
            for group, categories_members in groups.items():
                # Skip if the group is not found in the spreadsheet
                if group in [sheet.title for sheet in spreadsheet.worksheets()]:
                    """worksheet = spreadsheet.worksheet(group)
 
                    # Clear existing data in the worksheet (except row number 1)
                    clear_worksheet(worksheet)
                    """
 
                    # Prepare the data to be appended
                    append_data = []
                    for category, jira_group, wiki_group, members in categories_members:
                        append_data.append([category, jira_group, wiki_group, ", ".join(members)])
 
                    append_data_to_sheet(group, append_data)
                    """
                    # Append the data to the worksheet starting from row number 2
                    worksheet.append_rows(append_data, value_input_option='USER_ENTERED', insert_data_option='OVERWRITE')
                    # If all rows are full, clearing existing data fails on the next run, so add one row
                    worksheet.add_rows(1)"""
 
        else:
            print("API request failed.")
    else:
        print("API request failed with status code:", response.status_code)
 
 
# Helper function to append data to a specific sheet
def append_data_to_sheet(sheet_name, data):
    worksheet = None
 
    # Check if the sheet already exists
    if sheet_name in [sheet.title for sheet in spreadsheet.worksheets()]:
        worksheet = spreadsheet.worksheet(sheet_name)
    else:
        # If the sheet doesn't exist, create it
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="10")
 
    # If all rows are full, clearing existing data fails on the next run, so add rows
    worksheet.add_rows(1000)
     
    # Clear existing data in the worksheet (except row number 1)
    existing_data = worksheet.get_all_values()
    existing_data[1:] = []
    worksheet.clear()
    worksheet.update(existing_data)
 
    total_rows = worksheet.row_count
    print (total_rows)
 
    # Append the data to the worksheet starting from row number 2
    worksheet.append_rows(data, value_input_option='USER_ENTERED', insert_data_option='OVERWRITE')
     
    print(sheet_name)
    print(len(data))
    print(worksheet.row_count)
    # Delete empty rows
    requests = [
        {
            "deleteDimension": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "ROWS",
                    "startIndex": len(data) + 1,
                    "endIndex": total_rows - 1
                }
            }
        }
    ]
 
    # Send the batch update request to the API
    response = spreadsheet.batch_update({"requests": requests})
 
    # Check if the request was successful
    if 'replies' in response:
        print("Rows deleted successfully.")
    else:
        print("API request failed")
 
 
    time.sleep(5)
 
 
 
# Compare Jira groups and members, add missing members to jira_add_member
def compare_and_generate_jira_add_member():
    global jira_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    jira_add_member = []
 
    for group_name, members in jira_group_and_member_data:
        # Find the matching group in sorted_categories
        matching_group = next((group_data for group_data in sorted_categories if group_data[1] == group_name), None)
 
        if matching_group:
            # Check for missing members
            category, jira_group, wiki_group, existing_members = matching_group
            missing_members = set(existing_members) - set(members.split(", "))
            if missing_members:
                jira_add_member.append([jira_group, ", ".join(missing_members)])
                     
    # Add missing members to the Jira group
    for group, members in jira_add_member:
        for member in members.split(", "):
            try:
                jira.add_user_to_group(username=member, group_name=group)
            except requests.exceptions.HTTPError as e:
                if "user is already a member" in str(e):
                    print("User '{}' is already a member of group '{}'".format(member, group))
                    continue
                elif "'{}' does not exist".format(member) in str(e):
                    print("Cannot add user. '{}' does not exist".format(member))
                    continue
                else:
                    raise e
                         
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Member Add(Jira)", jira_add_member)
 
    return jira_add_member
 
# Compare Jira groups and generate data for jira_add_group
def compare_and_generate_jira_add_group():
    global jira_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    jira_add_group = []
 
    for category, jira_group, wiki_group, _ in sorted_categories:
        if not jira_group:
            continue
        # Find the matching group in jira_group_and_member_data
        matching_group = next((group_data for group_data in jira_group_and_member_data if group_data[0] == jira_group), None)
 
        if not matching_group:
            try:
                jira.create_group(jira_group)
                jira_add_group.append([jira_group])
                string_builder.Add(jira_group + ", ")
            except requests.exceptions.HTTPError as e:
                string_builder.Add("\n" + jira_group + "failed to add \n")
                if "A group or user with this name already exists" in str(e):
                    print("Group '{}' already exists".format(jira_group))
                    continue
                else:
                    raise e
     
                 
 
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Group Create(Jira)", jira_add_group)
 
    # Append string builder success flag
    string_builder.Add("\nSuccessfully added new jira groups\n")
 
    return jira_add_group
 
# Compare Jira groups and generate data for jira_delete_group
def compare_and_generate_jira_delete_group():
    global jira_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    jira_delete_group = []
 
    for group_name, _ in jira_group_and_member_data:
        # Find the matching group in sorted_categories
        matching_group = next((group_data for group_data in sorted_categories if group_data[1] == group_name), None)
 
        if not matching_group:
            jira_delete_group.append([group_name])
 
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Group Delete(Jira)", jira_delete_group)
 
    return jira_delete_group
 
# Compare Wiki groups and members, add missing members to wiki_add_member
def compare_and_generate_wiki_add_member():
    global wiki_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    wiki_add_member = []
 
    for group_name, members in wiki_group_and_member_data:
        # Find the matching group in sorted_categories
        matching_group = next((group_data for group_data in sorted_categories if group_data[2] == group_name), None)
 
        if matching_group:
            # Check for missing members
            category, jira_group, wiki_group, existing_members = matching_group
            missing_members = set(existing_members) - set(members.split(", "))
            if missing_members:
                wiki_add_member.append([wiki_group, ", ".join(missing_members)])
 
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Member Add(Wiki)", wiki_add_member)
     
    return wiki_add_member
 
 
# Compare Wiki groups and generate data for wiki_add_group
def compare_and_generate_wiki_add_group():
    global wiki_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    wiki_add_group = []
 
    for category, jira_group, wiki_group, _ in sorted_categories:
        # Find the matching group in wiki_group_and_member_data
        matching_group = next((group_data for group_data in wiki_group_and_member_data if group_data[0] == wiki_group), None)
 
        if not matching_group:
            wiki_add_group.append([wiki_group])
 
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Group Create(Wiki)", wiki_add_group)
     
    return wiki_add_group
 
# Compare Wiki groups and generate data for wiki_delete_group
def compare_and_generate_wiki_delete_group():
    global wiki_group_and_member_data
    global sorted_categories
 
    # Initialize lists to store the generated data
    wiki_delete_group = []
 
    for group_name, _ in wiki_group_and_member_data:
        # Find the matching group in sorted_categories
        matching_group = next((group_data for group_data in sorted_categories if group_data[2] == group_name), None)
 
        if not matching_group:
            wiki_delete_group.append([group_name])
 
    # Append the generated data to the corresponding sheet
    append_data_to_sheet("Group Delete(Wiki)", wiki_delete_group)
     
    return wiki_delete_group
 
 
async def add_wiki_group():
    global wiki_add_group_data
 
    # Launch the browser
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
 
    # Navigate to the login page
    driver.get(wiki_script_url)
 
    # Input the username and password
    username_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'os_username')))
    username_field.send_keys(wiki_username)
    password_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'os_password')))
    password_field.send_keys(wiki_password)
 
    # Click the login button
    login_button = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'loginButton')))
    login_button.click()
 
    # Wait for the new page to load
    time.sleep(5)
 
    # Input the password again and click the confirm button
    password_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'password')))
    password_field.send_keys(wiki_password)
    confirm_button = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'authenticateButton')))
    confirm_button.click()
 
    # Wait for the new page to load
    time.sleep(5)
 
    # Loop through each group in wiki_add_group and run the script for each one
    for group in wiki_add_group_data:
        group_name = group[0]
        if (group_name == ""):
            continue
        script = 'import com.atlassian.confluence.user.DefaultUserAccessor\nimport com.atlassian.confluence.user.UserAccessor\nUserAccessor userAccessor = ComponentLocator.getComponent(UserAccessor.class)\nDefaultUserAccessor defaultuserAccessor = ComponentLocator.getComponent(DefaultUserAccessor.class)\ndefaultuserAccessor.addGroup("{}")'.format(group_name)
        # Set the script text to the current group's script
        script_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ace_layer ace_gutter-layer ace_folding-enabled"]')))
        try:
            script_field.click()
            text_field = driver.switch_to.active_element
            text_field.send_keys(script)
        except Exception as e:
            print(f"Error executing script: {e}")
            continue
        script_run_btn = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'script-run-btn')))
        script_run_btn.click()
 
        # Wait for one of the result or error containers to appear
        visible_selector = None
        while not visible_selector:
            try:
                result_container = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'result-container')))
                error_container = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'error-container')))
                if result_container.is_displayed():
                    visible_selector = result_container
                    string_builder.Add(group_name + ", ")
                elif error_container.is_displayed():
                    visible_selector = error_container
                    string_builder.Add("\n" + group_name + " failed to add " + visible_selector.text + "\n")
            except:
                pass
        # Refresh the page for the next iteration
        driver.refresh()
 
    # Close the browser
    driver.quit()
 
    # Append string builder success flag
    string_builder.Add("Successfully added new wiki groups\n")
 
async def add_wiki_member():
    global wiki_add_member_data
 
    # Launch the browser
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
 
    # Navigate to the login page
    driver.get(wiki_script_url)
 
    # Input the username and password
    username_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'os_username')))
    username_field.send_keys(wiki_username)
    password_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'os_password')))
    password_field.send_keys(wiki_password)
 
    # Click the login button
    login_button = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'loginButton')))
    login_button.click()
 
    # Wait for the new page to load
    time.sleep(5)
 
    # Input the password again and click the confirm button
    password_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'password')))
    password_field.send_keys(wiki_password)
    confirm_button = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'authenticateButton')))
    confirm_button.click()
 
    # Wait for the new page to load
    time.sleep(5)
 
    # Loop through each group in wiki_add_member_data and run the script for each one
    for group_data in wiki_add_member_data:
        group_name, members = group_data[0], group_data[1]
 
        # Construct the script for the current group and members
        script = 'import com.atlassian.confluence.user.UserAccessor\nimport com.atlassian.sal.api.component.ComponentLocator\nUserAccessor userAccessor = ComponentLocator.getComponent(UserAccessor)\nString groupName = "{}"\nString members = "{}"\n\ndef group = userAccessor.getGroup(groupName)\nmembers.split(", ").each {{ member -> def user = userAccessor.getUserByName(member)\n    userAccessor.addMembership(group, user)\n}}'.format(group_name, members)
 
        # Set the script text to the current group's script
        script_field = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ace_layer ace_gutter-layer ace_folding-enabled"]')))
        try:
            script_field.click()
            text_field = driver.switch_to.active_element
            text_field.send_keys(script)
        except Exception as e:
            print(f"Error executing script: {e}")
            continue
        script_run_btn = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID,'script-run-btn')))
        script_run_btn.click()
 
        # Wait for one of the result or error containers to appear
        visible_selector = None
        while not visible_selector:
            try:
                result_container = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'result-container')))
                error_container = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'error-container')))
                if result_container.is_displayed():
                    visible_selector = result_container
                elif error_container.is_displayed():
                    visible_selector = error_container
            except:
                pass
        print("working on..." + group_name + " " + members)
        print(visible_selector.text)
        # Refresh the page for the next iteration
        driver.refresh()
 
    # Close the browser
    driver.quit()
 
 
# post result to slack
def post_result():
    slack_webhook_payload = { "channel":"$YOUR CHANNEL$", "message": string_builder }
    slack_webhook_url = '$YOUR API URL FOR SENDING SLACK MESSAGE$'
    requests.post(slack_webhook_url, data=slack_webhook_payload).json()
    string_builder.__init__()
 
 
# Execute the functions for google sheet
display_wiki_groups_and_members()
display_jira_groups_and_members()
fetch_and_append_groupware_data()
 
# Compare and generate data for the user service
jira_add_group_data = compare_and_generate_jira_add_group()
jira_add_member_data = compare_and_generate_jira_add_member()
#jira_delete_group_data = compare_and_generate_jira_delete_group()
 
wiki_add_group_data = compare_and_generate_wiki_add_group()
wiki_add_member_data = compare_and_generate_wiki_add_member()
#wiki_delete_group_data = compare_and_generate_wiki_delete_group()
 
asyncio.get_event_loop().run_until_complete(add_wiki_group())
asyncio.get_event_loop().run_until_complete(add_wiki_member())
post_result()
