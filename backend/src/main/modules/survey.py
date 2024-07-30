import os
import csv
import time
import json
import threading
import requests
from config import *
from common_config import *
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
    