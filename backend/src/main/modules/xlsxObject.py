from email.mime import base
import json
import pandas as pd
import pyexcel
import pymongo
from datetime import datetime

from config import *


class new_xlsxObject:
  def __init__(self, id):
    
    client = pymongo.MongoClient(connectionUrl)
    validationDB = client[databaseName]
    collection = validationDB[collectionName]
    query = {"id":id}
    result = collection.find(query)
    if result.count() == 1:
      for i in result:
        self.metadata = i
    else:
      print("Multiple id found for requested id::", id)
    sheetData = pyexcel.get_book_dict(file_name=self.metadata["xlsxPath"])
    self.xlsxData = {}
    for key,item in sheetData.items():
      if key not in self.metadata["basicValidation"]["skipSheet"]:
        self.xlsxData[key] = item
        self.xlsxData[key] = self.xlsxData[key][1:]
    self.dateFormat = "%d-%m-%Y"

  
  def getSheetNames(self):
    if self.metadata["xlsxPath"].split('.')[-1] != "xlsx":
      return "Invalid file"
    sheet = pyexcel.get_book_dict(file_name=self.metadata["xlsxPath"])
    return [key for key, item in sheet.items()]

  def checkSheetExists(self):
    if self.metadata["basicValidation"]["tabNames"]["tabNameList"] != self.getSheetNames():
      return False, self.metadata["basicValidation"]["tabNames"]["errMessage"]+str(set(self.metadata["basicValidation"]["tabNames"]["tabNameList"]).difference(set(self.getSheetNames()))),self.metadata["basicValidation"]["tabNames"]["suggestion"]
    return True
  
  def checkColumnsExists(self):
    columnData = {}
    try:
      for sheet in self.xlsxData:
        
        df = pd.DataFrame(self.xlsxData[sheet][1:], columns = self.xlsxData[sheet][0])
        
        for sheetColumn in self.metadata["basicValidation"]["reqColumnNames"]["reqColumnNamesList"]:
          columnData[sheetColumn["tabName"]] = sheetColumn["columName"]

        if len(set(columnData[sheet]).difference(set(df.columns.values))) > 0:
          print( columnData[sheet] , list(df.columns.values))
          return False, self.metadata["basicValidation"]["reqColumnNames"]["errMessage"]+sheet,self.metadata["basicValidation"]["reqColumnNames"]["suggestion"]
    except Exception as e:
      return False, e
    return True

  def checkDateFormat(self,testDate):
    return bool(datetime.strptime(testDate, self.dateFormat))

  def checkDates(self):
    try:
      baseSheet = self.metadata["advanceValidation"]["dateColumns"]["baseSheet"]
      startDateColumn = self.metadata["advanceValidation"]["dateColumns"][baseSheet]["startDateColumn"]
      endDateColumn = self.metadata["advanceValidation"]["dateColumns"][baseSheet]["endDateColumn"]
      # print(self.xlsxData[baseSheet][1])
      df = pd.DataFrame(self.xlsxData[baseSheet][1:], columns = self.xlsxData[baseSheet][0])
      
      self.startDate, self.endDate = df[startDateColumn][0], df[endDateColumn][0]
            
      # if not(isinstance(self.startDate, str) and isinstance(self.endDate, str)):
      #   return False,self.metadata["advanceValidation"]["dateColumns"]["errMessage"]+baseSheet,self.metadata["advanceValidation"]["dateColumns"]["suggestion"]


      if not self.checkDateFormat(self.startDate) or not self.checkDateFormat(self.endDate):
        return False,self.metadata["advanceValidation"]["dateColumns"]["errMessage"]+baseSheet,self.metadata["advanceValidation"]["dateColumns"]["suggestion"]

      for sheet in self.xlsxData:
        if sheet != baseSheet and sheet in self.metadata["advanceValidation"]["dateColumns"].keys():
          df = pd.DataFrame(self.xlsxData[sheet][1:], columns = self.xlsxData[sheet][0], dtype=str)
          
          for startDate, endDate in zip(df[self.metadata["advanceValidation"]["dateColumns"][sheet]["startDateColumn"]],df[self.metadata["advanceValidation"]["dateColumns"][sheet]["endDateColumn"]]):
            if not bool(datetime.strptime(startDate, self.dateFormat)) or not bool(datetime.strptime(endDate, self.dateFormat)):
              return False,self.metadata["advanceValidation"]["dateColumns"]["errMessage"]+sheet,self.metadata["advanceValidation"]["dateColumns"]["suggestion"]

            if not (self.startDate <= startDate and self.endDate >= endDate):
              return False,self.metadata["advanceValidation"]["dateColumns"]["errMessage"]+sheet,self.metadata["advanceValidation"]["dateColumns"]["suggestion"]
    
    except Exception as e:
      return False, e
        

# if all(x in list(df.columns.values) for x in columnData[sheet]):
      #   print( columnData[sheet] , list(df.columns.values))
      #   return False, self.metadata["basicValidation"]["reqColumnNames"]["errMessage"]+sheet,self.metadata["basicValidation"]["reqColumnNames"]["suggestion"]
      