from jira import JIRA
from openpyxl import Workbook  #output is xlsx files
from datetime import datetime

# update auth info and list of Jira keys to dump
jira_server_url = "https://<your instance here>.atlassian.net" 
jira_email = "email address of jira user here"
jira_api_token = "your token here"
jira_key_list = ["P42-12112","P42-12066","P42-11755"]  #locate tickts of interest in jira and export


# Dump ticket summary data to xls -----------------------------------------------
def write_xls_key_list(data_list):
    wb = Workbook()
    ws = wb.active

    headerrow = ["Key", "Points", "In Progress Count", "Finish Date", "NEW", "READY FOR WORK", "IN PROGRESS", "READY FOR TESTING", "IN TEST", "READY FOR REVIEW", "DONE", "BLOCKED"]
    ws.append(headerrow)

    for item in data_list:
        row = [
            item.get("Key", 0),
            item.get("Points", 0),
            item.get("In Progress Count", 0),
            item.get("Finish Date", 0),
            item.get("NEW", 0),
            item.get("READY FOR WORK", 0),
            item.get("IN PROGRESS", 0),
            item.get("READY FOR TESTING", 0),
            item.get("IN TESTING", 0),
            item.get("READY FOR REVIEW", 0),
            item.get("DONE", 0),
            item.get("BLOCKED", 0),
        ]
        ws.append(row)
    wb.save("dump_keys.xlsx")


# Dump each status history transaction to xlsx ----------------------------------------------
def write_xls_dump(data_list):
    wb = Workbook()
    ws = wb.active

    headerrow = ["Key", "Points", "DateTime","FromStatus","ToStatus","Duration"]
    ws.append(headerrow)

    for item in data_list:
        row = [
            item.get("Key", ""),
            item.get("Points", ""),
            item.get("DateTime", ""),
            item.get("FromStatus", ""),
            item.get("ToStatus", ""),
            item.get("Duration", ""),
        ]
        ws.append(row)
    wb.save("dump.xlsx")



# Connect to Jira Server ----------------------------------------
try:
    jira = JIRA(server=jira_server_url, basic_auth=(jira_email, jira_api_token))
except Exception as e:
    print(f"Error connecting to Jira: {e}")
    exit(1)


# Build Jira Issues and Extract Changelog list --------------------------------------
data_list = []
for jira_key in jira_key_list:
    issue = jira.issue(jira_key, expand='changelog')
    changelog = issue.changelog

    data_row = {
        "Key": issue.key,
        "Points": issue.fields.customfield_10027,
        "DateTime": issue.fields.created,
        "FromStatus": '',
        "ToStatus": "NEW",
        'Duration': 0
    }
    data_list.append(data_row)

    for history in changelog.histories:
        for item in history.items:
            if item.field == 'status':
                if (item.fromString != item.toString):  #skip if status not changing
                    data_row = {
                        "Key": issue.key,
                        "Points": issue.fields.customfield_10027,
                        "DateTime": history.created,
                        "FromStatus": item.fromString.upper(),
                        "ToStatus": item.toString.upper(),
                        'Duration': None
                    }

                    data_list.append(data_row)


# Calculate duration of each status event --------------------------------------
sorted_data_list = sorted(data_list, key=lambda x: (x["Key"],x["DateTime"]))

for item in sorted_data_list:
    if item["FromStatus"] != "":
        last_item['Duration'] = (datetime.strptime(item["DateTime"], '%Y-%m-%dT%H:%M:%S.%f%z') - \
                           datetime.strptime(last_item["DateTime"], '%Y-%m-%dT%H:%M:%S.%f%z')).total_seconds() / 3600.0
    
    last_item = item

# build the dump file
write_xls_dump(sorted_data_list)   


# Summarize the status change events per status for each Jira issue -----------------------------------------------
status_flow_full = []
status_flow = {"NEW":0, "IN PROGRESS":0, "READY FOR TESTING":0, "IN TESTING":0, "READY FOR REVIEW":0, "DONE":0, "BLOCKED":0}
in_progress_count = 0
finish_date = None

last_item_key = sorted_data_list[0]["Key"]
for item in sorted_data_list:
    if item["ToStatus"] == "IN PROGRESS":
        in_progress_count += 1
    if item["ToStatus"] == "DONE":
        finish_date = item["DateTime"]
    if item["Key"] != last_item_key:
        data_row = {
            "Key": last_item_key, 
            "Points": item["Points"],
            "In Progress Count": in_progress_count,
            "Finish Date": finish_date,
            "NEW": status_flow.get("NEW", 0), 
            "READY FOR WORK": status_flow.get("READY FOR WORK", 0),
            "IN PROGRESS": status_flow.get("IN PROGRESS", 0), 
            "READY FOR TESTING": status_flow.get("READY FOR TESTING", 0), 
            "IN TESTING": status_flow.get("IN TESTING", 0), 
            "READY FOR REVIEW": status_flow.get("READY FOR REVIEW", 0), 
            "DONE": status_flow.get("DONE", 0), 
            "BLOCKED": status_flow.get("BLOCKED", 0)
        }
        status_flow_full.append(data_row)

        last_item_key = item["Key"]
        in_progress_count = 0
        finish_date = None
        status_flow = {"NEW":0, "IN PROGRESS":0, "READY FOR TESTING":0, "IN TESTING":0, "READY FOR REVIEW":0, "DONE":0, "BLOCKED":0}

    if (item["Duration"] is not None):
        status_flow[item["ToStatus"]] = status_flow.get(item["ToStatus"],0) + item["Duration"]

data_row = {
    "Key": last_item_key, 
    "Points": status_flow.get("Points", 0),
    "In Progress Count": in_progress_count,
    "Finish Date": finish_date,
    "NEW": status_flow.get("NEW", 0), 
    "READY FOR WORK": status_flow.get("READY FOR WORK", 0),
    "IN PROGRESS": status_flow.get("IN PROGRESS", 0), 
    "READY FOR TESTING": status_flow.get("READY FOR TESTING", 0), 
    "IN TESTING": status_flow.get("IN TESTING", 0), 
    "READY FOR REVIEW": status_flow.get("READY FOR REVIEW", 0), 
    "DONE": status_flow.get("DONE", 0), 
    "BLOCKED": status_flow.get("BLOCKED", 0)
}
status_flow_full.append(data_row)

# build the summary xlsx file
write_xls_key_list(status_flow_full)
