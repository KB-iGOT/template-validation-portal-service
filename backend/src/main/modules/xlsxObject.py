import pandas as pd
import pymongo
import re, requests
from datetime import datetime
import numpy as np
import wget
import os

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
    
      
      for key in self.xlsxData.keys():
        if key in self.sheetNames:
          newHeader = self.xlsxData[key].iloc[0]
          self.xlsxData[key] = self.xlsxData[key][1:]
          self.xlsxData[key].columns = newHeader
      self.emailRegex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
      self.pdInfo = {}           #Program designer belong to this orgIds
      self.pmInfo = {}          #Program manager belong to this orgIds
      self.stateId = {}         #State ids of given states
      self.ccInfo = {}
      self.stateCodeList = []
    
    else:
      print("Multiple/No id found for requested id::", id)
      self.success = False

    
    

  def requiredTrue(self, conditionData, sheetName, columnName,responseData):
    if conditionData["required"]["isRequired"]:
      if columnName not in self.xlsxData[sheetName].columns:
        responseData["data"].append({"errCode":errBasic, "sheetName":sheetName,"columnName":columnName,"errMessage":conditionData["required"]["errMessage"].format(columnName),"suggestion":conditionData["required"]["suggestion"].format(columnName, sheetName)})
      else:
        df = self.xlsxData[sheetName][columnName].isnull()
        if df.values.any():
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[df == True].values).tolist(),"errMessage":conditionData["required"]["errMessage"].format(columnName),"suggestion":conditionData["required"]["suggestion"].format(columnName, sheetName)})
        elif len(self.xlsxData[sheetName]) == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["required"]["errMessage"].format(columnName),"suggestion":conditionData["required"]["suggestion"].format(columnName, sheetName)})
    return responseData


  def uniqueTrue(self, conditionData, sheetName, columnName, multipleRow,responseData):
    if not multipleRow:
      if self.xlsxData[sheetName].shape[0] > 1:
        if all(x["errMessage"] == conditionData["unique"]["errMessage"].format("") for x in responseData["data"]):
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":"", "rowNumber": list(range(2, self.xlsxData[sheetName].shape[0]+1)),"errMessage":conditionData["unique"]["errMessage2"].format(sheetName),"suggestion":conditionData["unique"]["suggestion2"].format(sheetName)})
    elif conditionData["unique"]["isUnique"]:
      if columnName in self.xlsxData[sheetName].columns:
        if not self.xlsxData[sheetName][columnName].is_unique:
          df = self.xlsxData[sheetName][columnName].duplicated(keep=False)
          if len(set((df.index[df == True].values).tolist()) - set(self.xlsxData[sheetName][columnName].loc[pd.isna(self.xlsxData[sheetName][columnName])].index.values.tolist())) == 0:
            return responseData
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":list(set((df.index[df == True].values).tolist()) - set(self.xlsxData[sheetName][columnName].loc[pd.isna(self.xlsxData[sheetName][columnName])].index.values.tolist())),"errMessage":conditionData["unique"]["errMessage"].format(columnName),"suggestion":conditionData["unique"]["suggestion"].format(columnName, sheetName)})
    return responseData


  def specialCharacters(self, conditionData, sheetName, columnName,responseData):
    if columnName in self.xlsxData[sheetName].columns:
      regexCompile = re.compile(str(conditionData["specialCharacters"]["notAllowedSpecialCharacters"]))

      df = self.xlsxData[sheetName][columnName].apply(lambda x: regexCompile.search(x))
      if not df.isnull().values.all():
        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[df.notnull()].values).tolist(),"errMessage":conditionData["specialCharacters"]["errMessage"].format(sheetName, columnName),"suggestion":conditionData["specialCharacters"]["suggestion"]})
      
    return responseData



  def specialCharacterName(self, conditionData, sheetName, columnName, responseData):
    if columnName in self.xlsxData[sheetName].columns:
      regexCompile = re.compile(str(conditionData["specialCharacterName"]["notAllowedSpecialCharacters"]))

      df = self.xlsxData[sheetName][columnName].apply(lambda x: regexCompile.search(x))
      if not df.isnull().values.all():
        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[df.notnull()].values).tolist(),"errMessage":conditionData["specialCharacterName"]["errMessage"].format(sheetName, columnName), "suggestion":conditionData["specialCharacterName"]["suggestion"]})
    return responseData

  def projectsSpecialCharacter(self, conditionData, sheetName, columnName, responseData):
    if columnName in self.xlsxData[sheetName].columns:
      regexCompile = re.compile(str(conditionData["projectsSpecialCharacter"]["notAllowedSpecialCharacters"]))

      df = self.xlsxData[sheetName][columnName].apply(lambda x: regexCompile.search(x))
      if not df.isnull().values.all():
        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[df.notnull()].values).tolist(),"errMessage":conditionData["projectsSpecialCharacter"]["errMessage"].format(sheetName, columnName), "suggestion":conditionData["projectsSpecialCharacter"]["suggestion"]})
    return responseData

  def recommendedForCheck(self, conditionData, sheetName, columnName, multipleRow,responseData):
    rolesList = []
    for roles in conditionData["recommendedForCheck"]["roles"]:
      rolesList.append(roles["code"])
    if len(rolesList) == 0:
      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":"recommendedFor role list is empty in the backend", "suggestion":"Please at least one role in the backend"})
      return responseData
    for idx, row in self.xlsxData[sheetName].iterrows():
      if idx > 1 and not multipleRow:
        break
      if row[columnName] == row[columnName]:
        df = [y.strip() for y in row[columnName].split(",")]
        if not all(item in rolesList for item in df):
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":conditionData["recommendedForCheck"]["errMessage"], "suggestion":conditionData["recommendedForCheck"]["suggestion"]})
                

    return responseData
  
  def dateFormatFun(self, conditionData, sheetName, columnName,responseData):
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
      
    return responseData
  
  

  def pdRoleCheck(self, conditionData, sheetName, columnName, newToken, multipleRow,responseData):
    self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
    self.xlsxData[sheetName][columnName] = self.xlsxData[sheetName][columnName].fillna("None")
    self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
    if columnName in self.xlsxData[sheetName].columns:  
      conditionData["pdRoleCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
      
      for index, row in self.xlsxData[sheetName].iterrows():
        if index > 1 and not multipleRow:
          break

        if row[columnName] == "None":
          continue
        
        conditionData["pdRoleCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

        if row["isEmail"] == None:
          conditionData["pdRoleCheck"]["body"]["request"]["filters"]["userName"] = conditionData["pdRoleCheck"]["body"]["request"]["filters"].pop("email")

        df = requests.post(url=hostUrl+conditionData["pdRoleCheck"]["api"],headers=conditionData["pdRoleCheck"]["headers"],json=conditionData["pdRoleCheck"]["body"])
        # print(df.json())

        if df.json()["result"]["response"]["count"] == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pdRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pdRoleCheck"]["suggestion"]})
        else:
          self.pdInfo[row[columnName]] = False
          for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
            if conditionData["pdRoleCheck"]["role"] in orgData["roles"]:
              self.pdInfo[row[columnName]] = True
              break
              # self.pdInfo[row[columnName]].update({orgData["orgName"]: orgData["organisationId"], "userName":df.json()["result"]["response"]["content"][0]["userName"]})
          if not self.pdInfo[row[columnName]]:
            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pdRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pdRoleCheck"]["suggestion"]})

    return responseData

    
  def stateCheck(self, conditionData, sheetName, columnName,responseData):
    # if columnName in self.xlsxData[sheetName].columns:
    if self.xlsxData[sheetName][columnName].iloc[0] == self.xlsxData[sheetName][columnName].iloc[0]:
      stateList = [item.strip() for item in self.xlsxData[sheetName][columnName].iloc[0].split(",")]
      for stateName in stateList:
        conditionData["stateCheck"]["body"]["request"]["filters"]["name"] = stateName
      
        df = requests.post(url=preprodHostUrl+conditionData["stateCheck"]["api"],headers=conditionData["stateCheck"]["headers"],json=conditionData["stateCheck"]["body"])
        
        if df.json()["result"]["count"] == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["stateCheck"]["errMessage"].format(stateName), "suggestion":conditionData["stateCheck"]["suggestion"]})
        else:
          self.stateCodeList.append(df.json()["result"]["response"][0]["code"])
          self.stateId[df.json()["result"]["response"][0]["id"]] = df.json()["result"]["response"][0]["name"]

    return responseData


  # def orgCheck(self, conditionData, sheetName, columnName, responseData):
  #   # if columnName in self.xlsxData[sheetName].columns:
  #   for orgId in self.stateId.keys():
  #     conditionData["orgCheck"]["body"]["request"]["filters"]["locationIds"] = orgId
    
  #     df = requests.post(url=preprodHostUrl+conditionData["orgCheck"]["api"],headers=conditionData["orgCheck"]["headers"],json=conditionData["orgCheck"]["body"])
  #     if df.json()["result"]["response"]["count"] == 0:
  #       responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["orgCheck"]["errMessage"].format(self.stateId[orgId]), "suggestion":conditionData["orgCheck"]["suggestion"]})
  #     else:
  #       self.stateOrgNames[df.json()["result"]["response"]["content"][0]["orgName"]] = df.json()["result"]["response"]["content"][0]["id"]

  #   return responseData

  def districtCheck(self, conditionData, sheetName, columnName,responseData):

    if columnName in self.xlsxData[sheetName].columns:  
      # if self.xlsxData[sheetName][columnName].iloc[0] == self.xlsxData[sheetName][columnName].iloc[0]:
        
      districtList = [item.strip() for item in self.xlsxData[sheetName][columnName].iloc[0].split(",")]
      for districtName in districtList:
        conditionData["districtCheck"]["body"]["request"]["filters"]["name"] = districtName
      
        df = requests.post(url=preprodHostUrl+conditionData["districtCheck"]["api"],headers=conditionData["districtCheck"]["headers"],json=conditionData["districtCheck"]["body"])
        
        if df.json()["result"]["count"] == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["districtCheck"]["errMessage"].format(districtName), "suggestion":conditionData["districtCheck"]["suggestion"]})
        else:
          if df.json()["result"]["response"][0]["parentId"] not in self.stateId.keys():
            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":1,"errMessage":conditionData["districtCheck"]["errMessage"].format(districtName), "suggestion":conditionData["districtCheck"]["suggestion"]})

    return responseData
  

  def pmRoleCheck(self, conditionData, sheetName, columnName, newToken, multipleRow,responseData):
    self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
    self.xlsxData[sheetName][columnName] = self.xlsxData[sheetName][columnName].fillna("None")
    self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
    if columnName in self.xlsxData[sheetName].columns:  
      conditionData["pmRoleCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
      
      for index, row in self.xlsxData[sheetName].iterrows():
        if index > 1 and not multipleRow:
          break
        if row[columnName] == "None":
          continue
                                  
        conditionData["pmRoleCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

        if row["isEmail"] == None:
          conditionData["pmRoleCheck"]["body"]["request"]["filters"]["userName"] = conditionData["pmRoleCheck"]["body"]["request"]["filters"].pop("email")

        df = requests.post(url=hostUrl+conditionData["pmRoleCheck"]["api"],headers=conditionData["pmRoleCheck"]["headers"],json=conditionData["pmRoleCheck"]["body"])
        
        if df.json()["result"]["response"]["count"] == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pmRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pmRoleCheck"]["suggestion"]})
        else:
          self.pmInfo[row[columnName]] = False
          for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
            if conditionData["pmRoleCheck"]["role"] in orgData["roles"]:
              self.pmInfo[row[columnName]] = True
              break
              # self.pmInfo[row[columnName]].update({orgData["orgName"]: orgData["organisationId"],"userName":df.json()["result"]["response"]["content"][0]["userName"]})
          if not self.pmInfo[row[columnName]]:
            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["pmRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["pmRoleCheck"]["suggestion"]})


    return responseData

  def ccRoleCheck(self, conditionData, sheetName, columnName, newToken, multipleRow,responseData):
    self.xlsxData[sheetName] = self.xlsxData[sheetName].drop(columns="isEmail", errors="ignore")
    self.xlsxData[sheetName][columnName] = self.xlsxData[sheetName][columnName].fillna("None")
    self.xlsxData[sheetName]["isEmail"] = self.xlsxData[sheetName][columnName].apply(lambda x: re.fullmatch(self.emailRegex, x))
    if columnName in self.xlsxData[sheetName].columns:  
      conditionData["ccRoleCheck"]["headers"]["X-authenticated-user-token"] = newToken.json()["access_token"]
      
      for index, row in self.xlsxData[sheetName].iterrows():
        if index > 1 and not multipleRow:
          break
        if row[columnName] == "None":
          continue
                                  
        conditionData["ccRoleCheck"]["body"]["request"]["filters"]["email"] = row[columnName]

        if row["isEmail"] == None:
          conditionData["ccRoleCheck"]["body"]["request"]["filters"]["userName"] = conditionData["ccRoleCheck"]["body"]["request"]["filters"].pop("email")

        df = requests.post(url=hostUrl+conditionData["ccRoleCheck"]["api"],headers=conditionData["ccRoleCheck"]["headers"],json=conditionData["ccRoleCheck"]["body"])
        
        if df.json()["result"]["response"]["count"] == 0:
          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["ccRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["ccRoleCheck"]["suggestion"]})
        else:
          self.ccInfo[row[columnName]] = False
          for orgData in df.json()["result"]["response"]["content"][0]["organisations"]:
            if conditionData["ccRoleCheck"]["role"] in orgData["roles"]:
              self.ccInfo[row[columnName]] = True
              break
              # self.ccInfo[row[columnName]].update({orgData["orgName"]: orgData["organisationId"],"userName":df.json()["result"]["response"]["content"][0]["userName"]})
          if not self.ccInfo[row[columnName]]:
            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":index,"errMessage":conditionData["ccRoleCheck"]["errMessage"].format(row[columnName]), "suggestion":conditionData["ccRoleCheck"]["suggestion"]})


    return responseData

  def storeResponse(self, conditionData, sheetName, columnName, multipleRow,responseData):
    self.response = {}
    for idx, row in self.xlsxData[sheetName].iterrows():
      if idx > 1 and not multipleRow:
        break
      self.response[row[columnName]] = {}
      for col in conditionData["storeResponse"]["columnNames"]:
        self.response[row[columnName]][col] = row[col]
    return responseData
  

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
      multipleRow = data["multipleRowsAllowed"]
      if sheetName in self.xlsxData.keys():
        for columnData in data["columns"]:
          columnName = columnData["name"]
          for conditionName in columnData["conditions"]:
            query = {"name": conditionName}
            result = collection.find(query)
            for conditionData in result:
            
              if conditionData["name"] == "requiredTrue":
                try:
                  responseData = self.requiredTrue(conditionData, sheetName, columnName,responseData)
                except Exception as e:
                  print(e, sheetName, columnName, "requiredTrue")
                  continue
              elif conditionData["name"] == "uniqueTrue":
                try:
                  responseData = self.uniqueTrue(conditionData, sheetName, columnName, multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"uniqueTrue")
                  continue  
              
              elif conditionData["name"] == "specialCharacters":
                try:
                  responseData = self.specialCharacters(conditionData, sheetName, columnName,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"specialCharacters")
                  continue

              elif conditionData["name"] == "specialCharacterName":
                try:
                  responseData = self.specialCharacterName(conditionData, sheetName, columnName,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"specialCharacterName")
                  continue
              
              elif conditionData["name"] == "projectsSpecialCharacter":
                try:
                  responseData = self.projectsSpecialCharacter(conditionData, sheetName, columnName,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"projectsSpecialCharacter")
                  continue
            
              elif conditionData["name"] == "dateFormat":
                try:
                  responseData = self.dateFormatFun(conditionData, sheetName, columnName,responseData)    
                except Exception as e:
                  print(e, sheetName, columnName,"dateFormat") 
                  continue
              
              elif conditionData["name"] == "stateCheck":
                try:   
                  responseData = self.stateCheck(conditionData, sheetName, columnName,responseData)             
                except Exception as e:
                  print(e, sheetName, columnName,"stateCheck")
                  continue
              
              # elif conditionData["name"] == "orgCheck":
              #   try:   
              #     responseData = self.orgCheck(conditionData, sheetName, columnName,responseData)             
              #   except Exception as e:
              #     print(e, sheetName, columnName,"orgCheck")
              #     continue

              elif conditionData["name"] == "districtCheck":
                try:   
                  responseData = self.districtCheck(conditionData, sheetName, columnName,responseData)             
                except Exception as e:
                  print(e, sheetName, columnName,"districtCheck")
                  continue
              
              elif conditionData["name"] == "pdRoleCheck":
                try:
                  responseData = self.pdRoleCheck(conditionData, sheetName, columnName, newToken, multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"pdRoleCheck")
                  continue
              
              elif conditionData["name"] == "pmRoleCheck":
                try:
                  responseData = self.pmRoleCheck(conditionData, sheetName, columnName, newToken, multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"pmRoleCheck")
                  continue
              
              elif conditionData["name"] == "ccRoleCheck":
                try:
                  responseData = self.ccRoleCheck(conditionData, sheetName, columnName, newToken, multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"ccRoleCheck")
                  continue
              
              
              elif conditionData["name"] == "recommendedForCheck":
                try:
                  responseData = self.recommendedForCheck(conditionData, sheetName, columnName,multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"recommendedForCheck")
                  continue
              
              elif conditionData["name"] == "storeResponse":
                try:
                  responseData = self.storeResponse(conditionData, sheetName, columnName,multipleRow,responseData)
                except Exception as e:
                  print(e, sheetName, columnName,"storeResponse")
                  continue

      else:
        if data["required"]:
          responseData["data"].append({"errCode":errBasic, "sheetName":sheetName, "columnName":"","errMessage":data["errMessage"].format(sheetName),"suggestion":data["suggestion"].format(sheetName)})
    
    return responseData

  def customCondition(self):
    responseData = {"data":[]} 
    
    for data in self.metadata["validations"]:
      sheetName = data["name"]
      multipleRow = data["multipleRowsAllowed"]

      if sheetName in self.xlsxData.keys():
        for columnData in data["columns"]:
          columnName = columnData["name"]
          try:
            if "customConditions" in columnData.keys():
              for customKey in columnData["customConditions"].keys():
                
                if customKey == "requiredValue":
                  for idx, row in self.xlsxData[sheetName].iterrows():
                    if idx > 1 and not multipleRow:
                      break
                    try:
                      dfTest = row[columnName].split(",")
                      for x in dfTest:
                        if x not in columnData["customConditions"]["requiredValue"]["values"]:
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":columnData["customConditions"]["requiredValue"]["errMessage"], "suggestion":(columnData["customConditions"]["requiredValue"]["suggestion"]).format(columnData["customConditions"]["requiredValue"]["values"])})
                    except Exception as e:
                      print(e,type(row[columnName]), row[columnName], "requiredValue")
                      continue
                  # df = self.xlsxData[sheetName][columnName].apply(lambda x: set([y.strip() for y in x.split(',')]).issubset(columnData["customConditions"]["requiredValue"]["values"]))
                  # if False in df.values:
                  #   responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[~df].values).tolist(),"errMessage":columnData["customConditions"]["requiredValue"]["errMessage"], "suggestion":(columnData["customConditions"]["requiredValue"]["suggestion"]).format(columnData["customConditions"]["requiredValue"]["values"])})
                
                
                elif customKey == "dependent":
                  for dependData in columnData["customConditions"][customKey]:
                    
                    if dependData["type"] == "operator":
                      try:
                        dateColumn = pd.to_datetime(self.xlsxData[sheetName][columnName], format='%d-%m-%Y')
                        baseDateColumn = pd.to_datetime(self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]], format='%d-%m-%Y')
                        
                        if dependData["dependsOn"]["dependentTabName"] == "Program Details" and sheetName != "Program Details":
                          baseDateColumn = pd.Series([baseDateColumn.iloc[0]]*dateColumn.size) 
                          baseDateColumn.index += 1
                        elif dateColumn.size != baseDateColumn.size:
                          print("Not allowed comparison", sheetName, columnName)
                          continue
                        if dependData["dependsOn"]["dependentColumnValue"] == ["<"]:
                          df = dateColumn <= baseDateColumn
                        elif dependData["dependsOn"]["dependentColumnValue"] == [">"]:
                          df = dateColumn >= baseDateColumn
                        
                        if False in df.values:
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[~df].values).tolist(),"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})
                      except Exception as e:
                        print(e, sheetName, "operator")
                        continue
                      # self.xlsxData[sheetName][columnName] = pd.to_datetime(self.xlsxData[sheetName][columnName], format='%d-%m-%Y')
                      # self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]] = pd.to_datetime(self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]], format='%d-%m-%Y')
                      
                      # if dependData["dependsOn"]["dependentColumnValue"] == ["<"]:
                      #   df = self.xlsxData[sheetName][columnName] <= self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]]
                      # elif dependData["dependsOn"]["dependentColumnValue"] == [">"]:
                      #   df = self.xlsxData[sheetName][columnName] >= self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]]
                      # print(sheetName, columnName)
                      # print(dependData["dependsOn"]["dependentTabName"],dependData["dependsOn"]["dependentColumnName"])
                      # print(df)

                      # if False in df.values:
                        # responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(df.index[~df].values).tolist(),"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})
                    
                    elif dependData["type"] == "attribute":
                      df = self.xlsxData[sheetName][columnName].str.split(",").apply(lambda x : [y.strip() for y in x])
                      attributeData = getattr(self,dependData["attributeName"])
                      count = 0
                        
                      for testList in df:
                        count += 1
                        if count > 1 and not multipleRow:
                          break
                  
                        for test in testList:
                          try:
                            if test in attributeData[self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]].iloc[count-1]].keys():
                              print("Allowed", sheetName, columnName,test)
                            elif "attributeKey" in dependData.keys():
                              if test == attributeData[self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]].iloc[count-1]][dependData["attributeKey"]]:
                                print("Allowed", sheetName, columnName,test)
                              else:
                                print("Not allowed", sheetName, columnName, test)
                                responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":dependData["errMessage"].format(test), "suggestion":dependData["suggestion"]})
                            else:
                              print("Not alllowed", sheetName, columnName,test)
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":dependData["errMessage"].format(test), "suggestion":dependData["suggestion"]})
                          except Exception as e:
                            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":dependData["errMessage"].format(testList), "suggestion":dependData["suggestion"]})
                            print(e, sheetName, columnName, "attribute")
                            continue

                        count += 1
                    
                    
                    elif dependData["type"] == "condition":
                      if dependData["conditionName"] == "subRoleCheck":
                        allowedSubRole = []
                        collection = self.validationDB[conditionCollection]
                        query = {"name": "subRoleCheck"}
                        result = collection.find(query)
                        for subRoleConfig in result:
                          for stateCode in self.stateCodeList:
                            subRoleConfig["subRoleCheck"]["body"]["request"]["subType"] = stateCode
                            subRoleData = requests.post(url=hostUrl+subRoleConfig["subRoleCheck"]["api"], headers=subRoleConfig["subRoleCheck"]["headers"], json=subRoleConfig["subRoleCheck"]["body"])
                            for z in subRoleData.json()["result"]["form"]["data"]["fields"][1]["children"]["administrator"][2]["templateOptions"]["options"]:
                              allowedSubRole.append(z["label"])
                              allowedSubRole.append(z["value"])
                              
                        
                        for idx, row in self.xlsxData[sheetName].iterrows():
                          if idx > 1 and not multipleRow:
                            break
                              
                          df = [y.strip() for y in row[dependData["dependsOn"]["dependentColumnName"]].split(",")]
                          if any(item in df for item in dependData["dependsOn"]["dependentColumnValue"]):
                            if row[columnName] != row[columnName]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"].format(row[columnName]), "suggestion":dependData["suggestion"]})
                            else:
                              dfTest = [y.strip() for y in row[columnName].split(",")]
                              if not all(x in allowedSubRole for x in dfTest):
                                responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"].format(dfTest), "suggestion":dependData["suggestion"]})


                    
                    elif dependData["type"] == "subset":
                      df = (self.xlsxData[dependData["dependsOn"]["dependentTabName"]][dependData["dependsOn"]["dependentColumnName"]].str.split(",")).apply(pd.Series).stack().unique().tolist()
                      df = [item.strip() for item in df]

                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if idx > 1 and not multipleRow:
                          break
                        dfTest = row[columnName].split(",")
                        if not all(x in df for x in dfTest):
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(df)})


                      # dfTest = (self.xlsxData[sheetName][columnName].str.split(","))
                      # dfTest = dfTest.apply(lambda x: all(item in x for item in df))
                      # if False in dfTest.values:
                      #   responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":(dfTest.index[~dfTest].values).tolist(),"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(df)})
                    
                    
                    elif dependData["type"] == "value":
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if idx > 1 and not multipleRow:
                          break
                        
                        if len(dependData["dependsOn"]["dependentColumnValue"]) == 0:
                          if row[columnName] == row[columnName]:
                            if row[dependData["dependsOn"]["dependentColumnName"]] == row[dependData["dependsOn"]["dependentColumnName"]]: 
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(dependData["dependsOn"]["dependentColumnValue"])})

                        elif dependData["dependsOn"]["dependentColumnValue"][0] == "*":
                          if row[columnName] == row[columnName]:
                            if row[dependData["dependsOn"]["dependentColumnName"]] != row[dependData["dependsOn"]["dependentColumnName"]]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(dependData["dependsOn"]["dependentColumnValue"])})  
                        else:
                          df = [y.strip() for y in row[dependData["dependsOn"]["dependentColumnName"]].split(",")]
                          if any(item in df for item in dependData["dependsOn"]["dependentColumnValue"]):
                            if row[columnName] != row[columnName] and dependData["isNeeded"]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(dependData["dependsOn"]["dependentColumnValue"])})
                          else:
                            if row[columnName] == row[columnName]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(dependData["dependsOn"]["dependentColumnValue"])})


                        # else:
                        #   if row[columnName] == row[columnName]:
                        #     responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"].format(dependData["dependsOn"]["dependentColumnValue"])})
                    elif dependData["type"] == "isInteger":
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if type(row[columnName]) == str:
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":(dependData["suggestion"]).format(dependData["range"])})
                        elif type(row[columnName]) == int:
                          if len(dependData["range"]) == 2:
                            if row[columnName] < dependData["range"][0] or row[columnName] > dependData["range"][1]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":(dependData["suggestion"]).format(dependData["range"])})
                    
                    elif dependData["type"] == "isParent":
                      parentTask = []
                      subTask = []
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if row[columnName] == row[columnName]:
                          if row[columnName] not in parentTask:
                            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})
                          else:
                            if subTask[parentTask.index(row[columnName])] == subTask[parentTask.index(row[columnName])]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})

                        parentTask.append(row[dependData["dependsOn"]["dependentColumnName"]])
                        subTask.append(row[columnName])
                    
                    elif dependData["type"] == "checkTask":
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if row[columnName] == row[columnName]:
                          if row["TaskId"] in self.xlsxData[sheetName][dependData["dependsOn"]["dependentColumnName"]].values.tolist():
                            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})

                    elif dependData["type"] == "checkResponse":
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if row[columnName] == row[columnName]:
                          if len(dependData["dependsOn"]["dependentColumnValue"]) != 0 :
                            # print(self.response[row[columnName]][dependData["dependsOn"]["dependentColumnName"]], dependData["dependsOn"]["dependentColumnValue"])
                            if self.response[row[columnName]][dependData["dependsOn"]["dependentColumnName"]] not in dependData["dependsOn"]["dependentColumnValue"]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":(dependData["suggestion"]).format(dependData["dependsOn"]["dependentColumnValue"])})
                          else:
                            for res in row[columnName].split(","):
                              try:
                                if self.response[row["parent_question_id"]][(dependData["dependsOn"]["dependentColumnName"]).format(res)] != self.response[row["parent_question_id"]][(dependData["dependsOn"]["dependentColumnName"]).format(res)]:
                                  responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":(dependData["errMessage"]).format(res), "suggestion":(dependData["suggestion"]).format(res)})
                              except Exception as e:
                                responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":(dependData["errMessage"]).format(res), "suggestion":(dependData["suggestion"]).format(res)})

                    elif dependData["type"] == "integerOperator":
                      for idx, row in self.xlsxData[sheetName].iterrows():
                        if row[columnName] == row[columnName]:
                          if dependData["dependsOn"]["dependentColumnValue"] == ["<"]:
                            if row[columnName] > row[dependData["dependsOn"]["dependentColumnName"]]:
                              responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":idx,"errMessage":dependData["errMessage"], "suggestion":dependData["suggestion"]})
                          

                elif customKey == "linkCheck":
                  count = 0
                  for x in self.xlsxData[sheetName][columnName]:
                    count += 1
                    if count > 1 and not multipleRow:
                      break
                    resourcePath = self.metadata["xlsxPath"].split(".")[0]+"_"+sheetName+"_"+str(count)+".xlsx"
                    if type(x) != str and x==x:
                      responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})
                      continue

                    if x[:39] == "https://docs.google.com/spreadsheets/d/":
                      x = x.split("/")[5]
                      x = "https://docs.google.com/spreadsheets/export?id={}&exportFormat=xlsx".format(x)
                      try:
                        wget.download(x, resourcePath)
                        if not os.path.exists(resourcePath):
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})
                        else:
                          os.remove(resourcePath)
                      except Exception as e:
                        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":str(count),"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})
                        print(e, sheetName, columnName,"linkCheck")
                        continue
                    else:
                      try:
                        x = requests.get("https://diksha.gov.in/api/content/v1/read/"+x.split("/")[-1].split("?")[0])
                        if x.json()["result"]["content"]["status"] != "Live":
                          responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})

                        if len(columnData["customConditions"][customKey]["allowedType"]) != 0:
                          if x.json()["result"]["content"]["contentType"] not in columnData["customConditions"][customKey]["allowedType"]:
                            responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})
                      except Exception as e:
                        responseData["data"].append({"errCode":errAdv, "sheetName":sheetName,"columnName":columnName,"rowNumber":count,"errMessage":columnData["customConditions"][customKey]["errMessage"], "suggestion":columnData["customConditions"][customKey]["suggestion"]})
                        print(e, sheetName, columnName,"linkCheck")
                        continue


                
          except Exception as e:
            print(e, sheetName,columnName)
            continue


    return responseData