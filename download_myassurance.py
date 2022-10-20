import requests
import json
import os
from datetime import datetime
import time
from tqdm import tqdm
import webbrowser
import urllib3

#Disable request warning 
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

#Set today's date
today = datetime.today().strftime('%d-%m-%Y')

#Define the download function
def download (url, line_id, type_id, headers, file_path, file_name):

	print("Downloading " + file_name)
	global payload_assurance
	payload_assurance['assuranceLineIds'] = line_id
	payload_assurance['assuranceTypeIds'] = type_id
	response = requests.post(url, data=json.dumps(payload_assurance), headers = headers, verify=False)
	if response.status_code == 200:
		print("Data Downloaded")
		create_file = os.path.join(file_path, file_name + " " + today + ".xlsx")
		with open (create_file, "wb") as file:
			file.write(response.content)
			print("File saved\n")
	else:
		print("Error Code: " + str(response.status_code) + "\n" + "Error Message: " + str(response.content) + "\n")

#Set timer to wait for other programs to load
print("Waiting for other programs to load...\n")

bar_width = [x for x in range(60)]

try:
	for i in tqdm(bar_width):
		time.sleep(1)
except KeyboardInterrupt:
	print('\nTimer Interrupted!')
	pass

print("\nRunning...\n\n")

#Define parameters
req_headers = {"Authorization": "TOKEN HERE",
                "Content-Type": "application/json",
                "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36", 
                "Cookie" : "_ga_1YF3ZRWTRK=GS1.1.1600055242.11.0.1600055242.0; _ga=GA1.2.799754803.1554275941; sso_1538981658268=VND_RSSO_V2.eyJpYXQiOjE2MTg4MDA3MTA5MjUsInNydiI6Imh0dHBzOi8vbG9naW4ucGljc3NjLnBldHJvbmFzLmNvbS9yc3NvIiwicmxtIjoiUmVtZWR5X1NBTUwiLCJ0b2tlbklkIjoiXzFlZjA2YWUxLTdkYzAtNDYzMS05MzRiLTQzMTVlNzEyMThmYSJ9; opuId=19debcb2-498d-470f-aa4b-f219ba898d10"
               }

with open ("headers.txt", 'rt') as myfile:
	token = myfile.read()

req_headers["Authorization"] = token

url_check = "https://myassurance.petronas.com/webapi/api/v1/myTasks/count"
url_summary = "https://myassurance.petronas.com/webapi/api/v1/reporting/extractAssurancePlanData"
url_checklist = "https://myassurance.petronas.com/webapi/api/v1/reporting/extractChecklistData"
url_finding = "https://myassurance.petronas.com/webapi/api/v1/reporting/extractFindingActionItemInfo"
url_time = "https://myassurance.petronas.com/webapi/api/v1/reportLog/getLastedUpdateDate"

payload_assurance = {
	"assuranceLineIds":"PUT YOUR ASSURANCE LINE ID HERE",
	"assuranceTypeIds":"PUT YOUR ASSURANCE TYPE ID HERE",
	"opuOrAssetIds":[],
	"assuranceCategory":"",
	"assesseeDepartmentIds":[],
	"assessorDepartmentIds":[],
	"startDateFrom":"",
	"startDateTo":"",
	"assuranceStatuses":"",
	"assuranceLevel":"",
	"assuranceStage":"",
	"riskAreaIds":[],
	"checklistIds":[],
	"assuranceYear":"",
	"isIncludePMSMapping":"false",
	"isRepetitiveFinding":"",
	"cosoComponent":[],
	"checklistTypeId":"",
	"elementIds":[],
	"assuranceProviders":[],
	"findingTypes":[],
	"findingClassifications":[],
	"rootCauses":[],
	"findingStatuses":[],
	"actionDueDateFrom":"",
	"actionDueDateTo":"",
	"actionPriorities":[],
	"actionCompletionStatuses":[],
	"actionCompletionDateFrom":"",
	"actionCompletionDateTo":"",
	"actionItemStatuses":[],
	"actionWorkflowStatuses":[],
	"oemsChecklistIds":[],
	"opuIds":[],
	"assetIds":[],
	"isSelectPetronasGroup":"false",
	"allSelectedEnterpriseIds":[]
	}

#Payload before myASSURANCE Update
'''payload_assurance = {
    "assuranceLineIds":["c760214e-74ed-406f-988d-ee4840958b47"],
    "assuranceTypeIds":"PUT YOUR ASSURANCE TYPE ID HERE",
    "opuOrAssetIds":["9e1e7ac2-0143-4c56-9b2d-78cf005a3503","19debcb2-498d-470f-aa4b-f219ba898d10","04627532-c79a-4a41-aa8b-8610b27f1d3c","413aa22f-9396-42a5-bfa1-d14438a03181","1b613476-b6e4-4bbe-b97d-177e29449da6","f5921dd8-1444-4c92-b279-f35ff573a0c7","79b8d17d-35df-4945-b70c-bc0005d9ec92"],
    "assuranceCategory":"",
    "assesseeDepartmentIds":[],
    "assessorDepartmentIds":[],
    "startDateFrom":"",
    "startDateTo":"",
    "assuranceStatuses":"",
    "assuranceLevel":"",
    "assuranceStage":"",
    "riskAreaIds":[],
    "assuranceYear":"",
    "opuIds":["19debcb2-498d-470f-aa4b-f219ba898d10"],
    "assetIds":["04627532-c79a-4a41-aa8b-8610b27f1d3c","413aa22f-9396-42a5-bfa1-d14438a03181","1b613476-b6e4-4bbe-b97d-177e29449da6","f5921dd8-1444-4c92-b279-f35ff573a0c7"],
    "isSelectPetronasGroup":'true'
    }'''

#Define file path
summary_path = r'C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Assurance\PRPC Data\LIVE\Assurance Data\\'
checklist_path = r'C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Assurance\PRPC Data\LIVE\Assurance Checklist Answers Data\\'
findings_path = r'C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Assurance\PRPC Data\LIVE\Findings and Action Items Data\\'
chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

#Define assurance line and type ID
FIRST_ID = ["c760214e-74ed-406f-988d-ee4840958b47"]
SECOND_ID = ["f5f4c0f4-7893-4253-97dc-ba7e62642f65"]
BLANK_ID = []
FA_ID = ["c4833d54-4a3b-4580-967b-83207bccfcf5"]
OEMS_ID = ["09b3d15d-924a-448f-92ae-3c7c0610d93a"]

#Define file name
FA_summary = "Extract Assurance Data"
FA_name = "Extract Assurance Checklist Answers Data"
OEMS_name = "OEMS Extract Assurance Checklist Answers Data"
FA_findings_name = "Extract Findings and Action Items Data"
OEMS_findings_name = "OEMS Extract Findings and Action Items Data"
second_line_findings_name = "Second Line Extract Findings and Action Items Data"

#Check on token validity
print("Checking access to myASSURANCE...")
check = requests.get(url_check, headers = req_headers, verify=False)
status = check.status_code
if status == 401:
	print('Access Revoked. Please go to %s and authorize access.' % "https://myassurance.petronas.com/")
	#webbrowser.get(chrome_path).open("https://myassurance.petronas.com", new=0, autoraise=True)
	webbrowser.open("https://myassurance.petronas.com", new=0, autoraise=True)
	new_token = input("Updated Token:")
	req_headers["Authorization"] = new_token

	with open ("headers.txt", 'w') as update_file:
		update_token = update_file.write(new_token)

	print("\n\nProceed to Download File\n")

elif status == 200:
	print("Access Granted\n")

else:
	print("Error occured. Status code: " + str(status_code))

#Check Latest Data Time
dt = requests.get(url_time, headers = req_headers, verify=False)
json_dt = dt.json()
online_date = json_dt[0]['latestUpdateTime']

print("myASSURANCE data as at: " + str(online_date) + "\n" )

with open ("datetime.txt", 'w') as update_date:
	write_date = update_date.write(online_date)

#FA Summary Download
#download(url_summary, FA_ID, req_headers, summary_path, FA_summary)

#FA Files Download
download(url_checklist, FIRST_ID, FA_ID, req_headers, checklist_path, FA_name)
download(url_finding, FIRST_ID, FA_ID, req_headers, findings_path, FA_findings_name)

#OEMS Files Download
download(url_checklist, FIRST_ID, OEMS_ID, req_headers, checklist_path, OEMS_name)
download(url_finding, FIRST_ID, OEMS_ID, req_headers, findings_path, OEMS_findings_name)

#Second Line Download
download(url_finding, SECOND_ID, BLANK_ID, req_headers, findings_path, second_line_findings_name)

print("Pushing update to Power BI...\n")

os.system('pause')