from cmath import nan
import email
from email.mime import base
from hashlib import new
import json
import pandas as pd
import pyexcel
import pymongo
import re, requests
from datetime import datetime
import numpy as np

from config import *

class xlsxObject:
  def __init__(self, id, xlsxPath):
    
    client = pymongo.MongoClient(connectionUrl)
    self.validationDB = client[databaseName]
    collection = self.validationDB[collectionName]
    query = {"id":id}
    
    result = collection.find(query)
    
    if result.count() == 1:
      self.success = True
      for i in result:
        self.metadata = i
      
      self.metadata["xlsxPath"] = xlsxPath
      if self.metadata["xlsxPath"].split('.')[-1] != "xlsx":
        raise AssertionError("Unexpected file format ")

      self.sheetNames = [sheetName["name"] for sheetName in self.metadata["validations"]]
      self.xlsxData = pd.read_excel(self.metadata["xlsxPath"], sheet_name=None)
    
      # for df in self.xlsxData.values():
      #   df.fillna("", inplace = True)

      for key in self.xlsxData.keys():
        if key in self.sheetNames:
          newHeader = self.xlsxData[key].iloc[0]
          self.xlsxData[key] = self.xlsxData[key][1:]
          self.xlsxData[key].columns = newHeader
      self.emailRegex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
      self.orgId = {}           #Program designer belong to this orgIds
      self.stateId = {}         #Xlsx file has this orgId
    
    else:
      print("Multiple/No id found for requested id::", id)
      self.success = False

    
    

  
  def checkSheetExists(self):
    for data in self.metadata["validations"]:
      sheetName = data["name"]
      if data["required"]:
        if not data["multipleRowsAllowed"]: 
          if self.xlsxData[sheetName].shape[0] > 1:
            return False, sheetName+" does not allow multiple row"
        if sheetName not in list(self.xlsxData.keys()):
          return False,data["errMesage"].format(sheetName), data["suggestion"].format(sheetName)
    return True 
  
  def basicCondition(self):
    responseData = {"data":[]}
    collection = self.validationDB[conditionCollection]
    query = {"name": "tokenConfig"}
    result = collection.find(query)
    for tokenConfig in result:
      newToken = requests.post(url=hostUrl+tokenConfig["tokenApi"], headers=tokenConfig["tokenHeader"], data=tokenConfig["tokenData"])

    collection = self.validationDB[conditionCollection]
    for data in self.metadata["validations"]:
      sheetName = data["name"]
      if sheetName in self.xlsxData.keys() and data["required"]:
        for columnData in data["columns"]:
          columnName = columnData["name"]
          for conditionName in columnData["conditions"]:
            query = {"name": conditionName}
            result = collection.find(query)
            for conditionData in result:
            
              if conditionData["name"] == "requiredTrue":
                if conditionData["required"]["isRequired"]:
                  if columnName not in self.xlsxData[sheetName].columns:
                    responseData["data"].append({"errCode":errBasic, "sheetName":sheetName,"columnName":columnName,"errMessage":conditionData["required"]["errMessage"].format(columnName),"suggestion":conditionData["required"]["suggestion"].format(columnName, sheetName)})
                    
              elif conditionData["name"] == "uniqueTrue":
                if conditionData["unique"]["isUnique"]:
                  if columnName in self.xlsxData[sheetName].columns:
                    if not self.xlsxData[sheetName][columnName].is_unique:
                      df = self.xlsxData[sheetName][columnName].duplicated(keep=False)
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":list(df.index[df == True].values),"errMessage":conditionData["unique"]["errMessage"].format(columnName),"suggestion":conditionData["unique"]["suggestion"].format(columnName, sheetName)})
                      
              
              elif conditionData["name"] == "specialCharacters":
                if columnName in self.xlsxData[sheetName].columns:
                  regexCompile = re.compile(str(conditionData["specialCharacters"]["notAllowedSpecialCharacters"]))
                  try:
                    df = self.xlsxData[sheetName][columnName].apply(lambda x: regexCompile.search(x))
                    if not df.isnull().values.all():
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":list(df.index[df.notnull()].values),"errMessage":conditionData["specialCharacters"]["errMessage"].format(sheetName, columnName),"suggestion":conditionData["specialCharacters"]["suggestion"]})
                      
                  except Exception as e:
                    print(e, columnName)
              elif conditionData["name"] == "specialCharacterName":
                if columnName in self.xlsxData[sheetName].columns:
                  regexCompile = re.compile(str(conditionData["specialCharacterName"]["notAllowedSpecialCharacters"]))
                  try:
                    df = self.xlsxData[sheetName][columnName].apply(lambda x: regexCompile.search(x))
                    if not df.isnull().values.all():
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":list(df.index[df.notnull()].values),"errMessage":conditionData["specialCharacterName"]["errMessage"].format(sheetName, columnName), "suggestion":conditionData["specialCharacterName"]["suggestion"]})
                      
                  except Exception as e:
                    print(e, columnName)
              
              elif conditionData["name"] == "dateFormat":
                if columnName in self.xlsxData[sheetName].columns:

                  if conditionData["dateFormat"]["format"] == "DD-MM-YYYY":
                    self.dateFormat = "%d-%m-%Y"
                  elif conditionData["dateFormat"]["format"] == "YYYY-MM-DD":
                    self.dateFormat = "%Y-%m-%d"
                  else:
                    responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"errMessage":conditionData["dateFormat"]["errMessage"].format(sheetName, columnName), "suggestion":conditionData["dateFormat"]["suggestion"]}) 
                    
                  df = pd.to_datetime(self.xlsxData[sheetName][columnName], format=self.dateFormat, errors='coerce')
                  if not df.notnull().all():
                    responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[df.isnull()].values).tolist(),"errMessage":conditionData["dateFormat"]["errMessage"].format(sheetName, columnName), "suggestion":conditionData["dateFormat"]["suggestion"]})
                    
              
              elif conditionData["name"] == "programUserCheck":
                self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
                self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
                
                if columnName in self.xlsxData[sheetName].columns:  
                  conditionData["programUserCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
                  for index, row in self.xlsxData[sheetName].iterrows():
                    conditionData["programUserCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

                    if row["isEmail"] == None:
                      conditionData["programUserCheck"]["body"]["request"]["filters"]["userName"] = conditionData["programUserCheck"]["body"]["request"]["filters"].pop("email")
                    
                    df = requests.post(url=hostUrl+conditionData["programUserCheck"]["api"],headers=conditionData["programUserCheck"]["headers"],json=conditionData["programUserCheck"]["body"])
                    if df.json()["result"]["response"]["count"] == 0:
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["programUserCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["programUserCheck"]["suggestion"]})
                    else:
                      for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
                        self.orgId[orgData["organisationId"]] = orgData["orgName"]

              elif conditionData["name"] == "stateCheck":
                
                if columnName in self.xlsxData[sheetName].columns:
                  if self.xlsxData[sheetName][columnName].iloc[0] == self.xlsxData[sheetName][columnName].iloc[0]:
                    stateList = [item.strip() for item in self.xlsxData[sheetName][columnName].iloc[0].split(",")]
                    for stateName in stateList:
                      conditionData["stateCheck"]["body"]["request"]["filters"]["name"] = stateName
                    
                      df = requests.post(url=preprodHostUrl+conditionData["stateCheck"]["api"],headers=conditionData["stateCheck"]["headers"],json=conditionData["stateCheck"]["body"])
                      
                      if df.json()["result"]["count"] == 0:
                        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["stateCheck"]["errMessage"].format(stateName), "suggestion":conditionData["stateCheck"]["suggestion"]})
                      else:
                        self.stateId[df.json()["result"]["response"][0]["id"]] = df.json()["result"]["response"][0]["name"]

              elif conditionData["name"] == "districtCheck":
                
                if columnName in self.xlsxData[sheetName].columns:  

                  if self.xlsxData[sheetName][columnName].iloc[0] == self.xlsxData[sheetName][columnName].iloc[0]:
                    
                    districtList = [item.strip() for item in self.xlsxData[sheetName][columnName].iloc[0].split(",")]
                    for districtName in districtList:
                      conditionData["districtCheck"]["body"]["request"]["filters"]["name"] = districtName
                    
                      df = requests.post(url=preprodHostUrl+conditionData["districtCheck"]["api"],headers=conditionData["districtCheck"]["headers"],json=conditionData["districtCheck"]["body"])
                      
                      if df.json()["result"]["count"] == 0:
                        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["districtCheck"]["errMessage"].format(districtName), "suggestion":conditionData["districtCheck"]["suggestion"]})
                      else:
                        if df.json()["result"]["response"][0]["parentId"] not in self.stateId.keys():
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["districtCheck"]["errMessage"].format(districtName), "suggestion":conditionData["districtCheck"]["suggestion"]})

              elif conditionData["name"] == "userCheck":
                self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
                self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
                
                if columnName in self.xlsxData[sheetName].columns:  
                  conditionData["userCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
                  
                  for index, row in self.xlsxData[sheetName].iterrows():
                    conditionData["userCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

                    if row["isEmail"] == None:
                      conditionData["userCheck"]["body"]["request"]["filters"]["userName"] = conditionData["userCheck"]["body"]["request"]["filters"].pop("email")


                    df = requests.post(url=hostUrl+conditionData["userCheck"]["api"],headers=conditionData["userCheck"]["headers"],json=conditionData["userCheck"]["body"])
                    if df.json()["result"]["response"]["count"] == 0 :
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["userCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["userCheck"]["suggestion"]})
                    else:
                      orgList = []
                      for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
                        orgList.append(orgData["organisationId"])  
                      if not any(item in orgList for item in self.orgId.keys()):
                        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["userCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["userCheck"]["suggestion"]})
              
              
              elif conditionData["name"] == "pdRoleCheck":
                self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
                self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
                if columnName in self.xlsxData[sheetName].columns:  
                  conditionData["pdRoleCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
                  
                  for index, row in self.xlsxData[sheetName].iterrows():
                    conditionData["pdRoleCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

                    if row["isEmail"] == None:
                      conditionData["pdRoleCheck"]["body"]["request"]["filters"]["userName"] = conditionData["pdRoleCheck"]["body"]["request"]["filters"].pop("email")

                    df = requests.post(url=hostUrl+conditionData["pdRoleCheck"]["api"],headers=conditionData["pdRoleCheck"]["headers"],json=conditionData["pdRoleCheck"]["body"])
                    if df.json()["result"]["response"]["count"] > 0:
                      for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
                        if conditionData["pdRoleCheck"]["role"] not in orgData["roles"]:
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pdRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pdRoleCheck"]["suggestion"]})
                    else:
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pdRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pdRoleCheck"]["suggestion"]})
              
              elif conditionData["name"] == "pmRoleCheck":
                self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
                self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
                if columnName in self.xlsxData[sheetName].columns:  
                  conditionData["pmRoleCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
                  
                  for index, row in self.xlsxData[sheetName].iterrows():
                    conditionData["pmRoleCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

                    if row["isEmail"] == None:
                      conditionData["pmRoleCheck"]["body"]["request"]["filters"]["userName"] = conditionData["pmRoleCheck"]["body"]["request"]["filters"].pop("email")

                    df = requests.post(url=hostUrl+conditionData["pmRoleCheck"]["api"],headers=conditionData["pmRoleCheck"]["headers"],json=conditionData["pmRoleCheck"]["body"])
                    if df.json()["result"]["response"]["count"] > 0:
                    
                      for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
                        if conditionData["pmRoleCheck"]["role"] not in orgData["roles"]:
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pmRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pmRoleCheck"]["suggestion"]})
                    else:
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pmRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pmRoleCheck"]["suggestion"]})
      else:
        responseData["data"].append({"errCode":errBasic, "sheetName":sheetName,"errMessage":data["errMessage"].format(sheetName),"suggestion":data["suggestion"].format(sheetName)})
    return responseData

  def customCondition(self):
    responseData = {"data":[]} 
    
    for data in self.metadata["validations"]:
      sheetName = data["name"]
      if sheetName in self.xlsxData.keys():
        for columnData in data["columns"]:
          columnName = columnData["name"]
          
          if "customConditions" in columnData.keys():
            for customKey in columnData["customConditions"].keys():
              if customKey == "requiredValue":
                df = self.xlsxData[sheetName][columnName].apply(lambda x: set([y.strip() for y in x.split(', ')]).issubset(columnData["customConditions"]["requiredValue"]["values"]))
                if False in df.values:
                  responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[~df].values).tolist(),"errMessage":columnData["customConditions"]["requiredValue"]["errMessage"], "suggestion":(columnData["customConditions"]["requiredValue"]["suggestion"]).format(columnData["customConditions"]["requiredValue"]["values"])})
    return responseData