import os
import csv
import time
import json
import threading
import requests
from config import *
from common_config import *

class surveyCreate:
    def __init__(self):
        pass

    def generate_access_token(self):
        header_keyclock_user = {'Content-Type': keyclockapicontent_type}
        try:
            response = requests.post(
                url=host + keyclockapiurl,
                headers=header_keyclock_user,
                data=keyclockapibody
            )
            response.raise_for_status()
            access_token = response.json().get('access_token')
            if not access_token:
                raise ValueError("Access token not found in the response.")
            print("---> Access Token Generated!")
            return access_token
        except (requests.RequestException, ValueError) as e:
            print(f"Error generating access token: {e}")
            return None

    def fetch_solution_id(self, access_token, csv_file_path='solutions.csv'):
        print(access_token)
        if not access_token:
            print("Invalid access token.")
            return None

        solution_update_api = f"{internal_kong_ip_core}{dbfindapi_url}solutions"
        print("solutionUpdateApi:", solution_update_api)
        headers = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-authenticated-user-token': access_token,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }

        payload = {
            "query": {"status": "active"},
            "mongoIdKeys": ["_id", "solutionId", "metaInformation.solutionId"],"limit":1000
        }

        try:
            response = requests.post(
                url=solution_update_api,
                headers=headers,
                data=json.dumps(payload)
            )
            response.raise_for_status()
            result = response.json().get('result', [])
        except requests.RequestException as e:
            print(f"Error fetching solutions: {e}")
            return None

        file_exists = os.path.isfile(csv_file_path)
        with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['SOLUTION_ID', 'SOLUTION_NAME', 'SOLUTION_CREATED_DATE', 'STARTDATE', 'ENDDATE']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for item in result:
                solution_id = item.get('_id', 'N/A')
                solution_name = item.get('name', 'N/A')
                solution_createdat = item.get('createdAt', 'N/A')
                startdate = item.get('startDate', 'None')
                endate = item.get('endDate', 'None')

                writer.writerow({
                    'SOLUTION_ID': solution_id,'SOLUTION_NAME': solution_name,'SOLUTION_CREATED_DATE': solution_createdat,'STARTDATE': startdate,'ENDDATE': endate})

        print("Data written to CSV successfully.")
        self.schedule_deletion(csv_file_path)
        couldPathForCsv=self.uploadSuccessSheetToBucket(csv_file_path,access_token)
        print(couldPathForCsv,"couldPathForCsv")
        return couldPathForCsv
    
    def uploadSuccessSheetToBucket(self,csv_file_path,access_token):
        print("successSheetName----------",csv_file_path)
        persignedUrl = public_url_for_core_service + getpresignedurl
        solutionDump = "surveydump"
        
        presignedUrlBody = {
            "request": {
                solutionDump :{
                 
                    "files": [
                        csv_file_path
                    ]
            }
                
            },
            "ref": "solutionDump"
        }
        headerPreSignedUrl = {'Authorization': authorization,
                                   'X-authenticated-user-token': access_token,
                                   'Content-Type': content_type}
        responseForPresignedUrl = requests.request("POST", persignedUrl, headers=headerPreSignedUrl,
                                                    data=json.dumps(presignedUrlBody))
        
        if responseForPresignedUrl.status_code == 200:
            presignedResponse = responseForPresignedUrl.json()
            programupdateData = presignedResponse['result']
            fileUploadUrl = presignedResponse['result'][solutionDump]['files'][0]['url']
            if '?file=' in fileUploadUrl:
                downloadedurl = fileUploadUrl.split('?file=')[1]
            else:
                downloadedurl = None
            print(downloadSuccessSheet+downloadedurl,"click here to download")

            print(fileUploadUrl,"fileUploadUrlfileUploadUrl")
            print(presignedResponse['result'][solutionDump]['files'][0],"till here reacher")
            print("fileUploadUrl-------",fileUploadUrl)
            headers = {
                'Authorization': authorization,
                'X-authenticated-user-token': access_token,

            }

            files={
                'file': open(csv_file_path, 'rb')
            }
            print("files---------",files)

            response = requests.post(url=fileUploadUrl, headers=headers, files=files)
            if response.status_code == 200:
                print("File Uploaded successfully")
        print(downloadSuccessSheet+downloadedurl,"return downloadSuccessSheet+downloadedurl")
        return downloadSuccessSheet+downloadedurl
        
    def schedule_deletion(self,file_path):
        def delete_file():
            try:
                time.sleep(60)
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"File {file_path} deleted successfully.")
                else:
                    print(f"File {file_path} not found.")
            except Exception as e:
                print(f"Error deleting file: {e}")

        threading.Thread(target=delete_file, daemon=True).start()

