import os
import csv
import time
import json
import threading
import requests
from backend.src.main.modules.config import *
# from common_config import *
from datetime import datetime
from requests import get,post

class SurveyCreate:
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

            return access_token
        except (requests.RequestException, ValueError) as e:
            return None

    def fetch_solution_id(self, access_token, resurceType):
        if not access_token:
            return None
        solution_update_api = f"{internal_kong_ip_core}{dbfindapi_url}solutions"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-authenticated-user-token': access_token,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }

        payload = {
            "query": {"status": "active"},
            "resourceType": [resurceType + " Solution"],
            "mongoIdKeys": ["_id", "solutionId", "metaInformation.solutionId"],
            "limit": 1000
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
            return None
        
        all_solution_ids = {item['_id'] for item in result}
        all_parent_solution_ids = {item.get('parentSolutionId') for item in result if 'parentSolutionId' in item}

        solutions_data = []
        for item in result:
            solution_id = item.get('_id', 'N/A')
            parent_solution_id = item.get('parentSolutionId', 'N/A')
            if solution_id in all_parent_solution_ids:
                continue
            solution_data = {
                'SOLUTION_ID': solution_id,
                'SOLUTION_NAME': item.get('name', 'N/A'),
                'SOLUTION_CREATED_DATE': item.get('createdAt', 'N/A'),
                'START_DATE': item.get('startDate', 'None'),
                'END_DATE': item.get('endDate', 'None')
            }
            solutions_data.append(solution_data)

        solutions_data.sort(key=lambda x: datetime.strptime(x['SOLUTION_CREATED_DATE'], "%Y-%m-%dT%H:%M:%S.%fZ"), reverse=True)
        
        return solutions_data   

    def fetch_solution_id_csv(self, access_token, resurceType,csv_file_path='solutions.csv'):
        if not access_token:
            return None
        solution_update_api = f"{internal_kong_ip_core}{dbfindapi_url}solutions"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': authorization,
            'X-authenticated-user-token': access_token,
            'X-Channel-id': x_channel_id,
            'internal-access-token': internal_access_token
        }

        payload = {
            "query": {"status": "active"},
            "resourceType": [resurceType + " Solution"],
            "mongoIdKeys": ["_id", "solutionId", "metaInformation.solutionId"],
            "limit": 1000
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
            return None
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

        result.sort(key=lambda x: x.get('createdAt', 'N/A'), reverse=True)
        
        all_solution_ids = {item['_id'] for item in result}
        all_parent_solution_ids = {item['parentSolutionId'] for item in result if 'parentSolutionId' in item}

        file_exists = os.path.isfile(csv_file_path)
        with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['SOLUTION_ID', 'SOLUTION_NAME', 'SOLUTION_CREATED_DATE', 'START_DATE', 'END_DATE']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for item in result:
                solution_id = item.get('_id', 'N/A')
                parent_solution_id = item.get('parentSolutionId', 'N/A')
                if solution_id in all_parent_solution_ids:
                    continue
                solution_id = item.get('_id', 'N/A')
                solution_name = item.get('name', 'N/A')
                solution_createdat = item.get('createdAt', 'N/A')
                startdate = item.get('startDate', 'None')
                endate = item.get('endDate', 'None')

                writer.writerow({
                    'SOLUTION_ID': solution_id,'SOLUTION_NAME': solution_name,'SOLUTION_CREATED_DATE': solution_createdat,'START_DATE': startdate,'END_DATE': endate})

        print("Data written to CSV successfully.")
        self.schedule_deletion(csv_file_path)
        couldPathForCsv=self.uploadSuccessSheetToBucket(csv_file_path,access_token)
        return couldPathForCsv
    
    def uploadSuccessSheetToBucket(self,csv_file_path,access_token):
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
            # print(programupdateData)
            fileUploadUrl = presignedResponse['result'][solutionDump]['files'][0]['url'][0]
            downloadedurl = presignedResponse['result'][solutionDump]['files'][0]['getDownloadableUrl'][0]
            # print(downloadedurl,"downloadedurl")
            # print(fileUploadUrl)
            headers = {
                "Content-Type":"multipart/form-data"

            }
            files={
                'file': open(csv_file_path, 'rb')
            }
            response = requests.put(url=fileUploadUrl, headers=headers, files=files)
            print(response.status_code)
            if response.status_code == 200:
                print("File Uploaded successfully")
        return downloadedurl
        
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

        
        
    