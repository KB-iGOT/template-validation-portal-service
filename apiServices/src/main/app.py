#write your code
from flask import Flask, request , send_from_directory
import os,time,sys
from dotenv import load_dotenv
import json 
import hashlib 
import jwt
from flask_cors import CORS
import numpy as np
import datetime
import pymongo
import pandas as pd
import shutil
import openpyxl
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill


sys.path.append('../../..')
sys.path.append('../../../backend/src/main/modules/')
from backend.src.main.modules.xlsxObject import xlsxObject


def myconverter(obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, datetime.datetime):
            return obj.__str__()


STATIC_PATH = os.path.join(os.getcwd(),"tmp")
app = Flask(__name__,static_url_path="/tmp/")
CORS(app)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(BASE_DIR, '.env')  # just an e.g

if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
else:
    import sys
    print('".env" is missing.')
    sys.exit(1)

def connectDb(url,db,collection):
    client = pymongo.MongoClient(url)
    db = client[db]
    collectionData = db[collection]
    # print("Connection Status : ",client.server_info())
    return collectionData

def addComments(templatePath, errResponse):
    xlsxData = pd.read_excel(templatePath, sheet_name=None)
    errPath = templatePath.split(".")[0]+"_errFile"+".xlsx"
    shutil.copyfile(templatePath, errPath)

    workBook = load_workbook(errPath, data_only = True)
    try:
        for key in xlsxData.keys():
            newHeader = xlsxData[key].iloc[0]
            xlsxData[key] = xlsxData[key][1:]
            xlsxData[key].columns = newHeader
    except Exception as e:
        workBook.save(errPath)
        errResponse["result"]["errFileLink"] = "http://34.143.225.1/template/api/v1/errDownload?templatePath="+errPath
        return errResponse
    for result in errResponse["result"]:
        for errData in errResponse["result"][result]["data"]:
            if errData["columnName"] != "":
                try:
                    spreadSheet = workBook[errData["sheetName"]]
                except Exception as e:
                    print(e, errData["sheetName"]," sheet is missing")
                    continue
                try:
                    columnNumber = xlsxData[errData["sheetName"]].columns.get_loc(errData["columnName"])
                except Exception as e:
                    if spreadSheet.cell(2,1).comment is None:
                        spreadSheet.cell(2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                        spreadSheet.cell(2,1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    else:
                        spreadSheet.cell(2,1).comment=Comment(spreadSheet.cell(row=2, column=1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                        spreadSheet.cell(2,1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    continue
                if type(errData["rowNumber"]) is list:
                    for rowIndex in errData["rowNumber"]:
                        if spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).comment is None:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                            spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                        else:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).comment=Comment(spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                            spreadSheet.cell(row=rowIndex+2, column=columnNumber+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                elif type(errData["rowNumber"]) is int:
                    if spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).comment is None:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    else:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).comment=Comment(spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumber+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid") 
            else:
                if errData["errCode"] == 300:
                    workBook.create_sheet(errData["sheetName"])
                    spreadSheet = workBook[errData["sheetName"]]
                    spreadSheet.cell(2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                    spreadSheet.cell(2,1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    continue
                else:
                    spreadSheet = workBook[errData["sheetName"]]
                    for rowIndex in errData["rowNumber"]:
                        if spreadSheet.cell(rowIndex+2,1).comment is None: 
                            spreadSheet.cell(rowIndex+2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                            spreadSheet.cell(rowIndex+2,1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                        else:
                            spreadSheet.cell(rowIndex+2,1).comment=Comment(spreadSheet.cell(rowIndex+2,1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                            spreadSheet.cell(rowIndex+2,1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    continue

                    
    workBook.save(errPath)
    errResponse["result"]["errFileLink"] = "http://34.143.225.1/template/api/v1/errDownload?templatePath="+errPath
    return errResponse

@app.route("/template/api/v1/authenticate", methods = ['POST'])
def login():
    req_body = request.get_json()
    savedUSERNAME = os.environ.get('email')
    savedPASSWORD = os.environ.get('password')
    try:
        userName = req_body['request']['email']
        password = hashlib.md5(req_body['request']['password'].encode('utf-8'))

        client = pymongo.MongoClient(os.environ.get('mongoURL'))

        db = client[os.environ.get('db')]
        
        usersCollection = db[os.environ.get('userCollection')]
        
        users = usersCollection.count_documents({'userName' : userName , "password" : str(password.hexdigest())})

        if(users):
            # Exipry and other details can be added here
            message = {
                'iss': '',
                'email': userName
                }
            signing_key = os.environ.get("SECRET_KEY")
            encoded_jwt = jwt.encode({'message': message}, signing_key, algorithm='HS256')


            return {"status" : 200,"code" : "Authenticated","errorFlag" : False,"error" : [],"response" : {
                "accessToken" : encoded_jwt
            }}
        else:
            return {"status" : 404,"code" : "Error","errorFlag" : True,"error" : ["Username / Password Doesn't Match"],"response" : {
                "accessToken" : "" }}
    except Exception as e:
        return {"status" : 500,"code" : str(e) ,"errorFlag" : True,"error" : ["Error in reaching server"],"response" : {
                "accessToken" : "" }}

@app.route("/template/api/v1/signup", methods = ['POST'])
def signup():
    req_body = request.get_json()
    auth = request.headers.get('admin-token')
    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : {"templateLinks" : ""}}
    else:
        if not auth == os.environ.get('admin-token'):
            return {"status" : 500,"code" : "Not Authorized" , "result" : {"templateLinks" : ""}}

    try:
        userName = req_body['request']['email']
        password = hashlib.md5(req_body['request']['password'].encode('utf-8'))

        client = pymongo.MongoClient(os.environ.get('mongoURL'))

        db = client[os.environ.get('db')]
        
        usersCollection = db[os.environ.get('userCollection')]
        now = datetime.datetime.now()
        users = usersCollection.count_documents({'userName' : userName})
        
        if(users <= 0):
            users = usersCollection.insert_one({'userName' : userName , "password" : str(password.hexdigest()),"status" : "active","role" : "admin","createdAt" : str(now),"updatedAt" : str(now),"createdBy" : "admin"})
            return {"status" : 200,"code" : "Authenticated","errorFlag" : False,"error" : [],"response" : "User created Successfully."}
        else:
            return {"status" : 404,"code" : "Error","errorFlag" : True,"error" : ["UserName already exisiting."],"response" : {"accessToken" : "" }}
    except Exception as e:
        return {"status" : 500,"code" : str(e) ,"errorFlag" : True,"error" : ["Error in reaching server"],"response" : {"accessToken" : "" }}


@app.route("/template/api/v1/download/sampleTemplate", methods = ['GET'])
def sample():
    templateList = os.environ.get('templateList').split(",")
    templateListResp = []
    tem = os.environ.get('templateIds')
    tem = json.loads(tem)
 
    for i in templateList:
        templateListResp.append({"templateName" : i, "templateLink" : os.environ.get(i) , "templateCode" : tem[i]})

    return {"status" : 200,"code" : "OK" , "result" : {"templateLinks" : templateListResp}}

@app.route("/template/api/v1/upload", methods = ['POST'])
def upload():

    # Token validation
    auth = request.headers.get('Authorization')
    signing_key = os.environ.get("SECRET_KEY")
    payload = False
    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : {"templateLinks" : ""}}
    else:
        try:
            payload = jwt.decode(auth, signing_key, algorithms=['HS256'])
        except Exception as e:
            print(e)

    if(not payload):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : {"templateLinks" : "True"}}
    
    ALLOWED_EXTENSIONS = set(['xlsx'])

    if not os.path.exists(STATIC_PATH):
        os.makedirs(STATIC_PATH)

    if request.method == 'POST':
        # check if the post request has the file part

        if 'file' not in request.files:
            return {"status" : 500,"code" : "Required key missing!" , "result" : {"templateLinks" : ""}}
        file = request.files['file']

        ext = file.filename.split('.')
        if file and ext[1] in ALLOWED_EXTENSIONS:
            filename = file.filename
            #fileName clearing.
            filename = filename.replace(" ","_")
            filenameArr = filename.split(".")

            # ts stores the time in seconds
            ts = str(time.time()).replace(".","-")
            finalFileName = str(filenameArr[0])+str(ts)+"."+str(filenameArr[1])
            try:
                file.save(os.path.join(STATIC_PATH, finalFileName))
                # print(os.path.join(STATIC_PATH, finalFileName))

            except Exception as e:
                print(e)
                return {"status" : 500,"code" : "Server Error" , "result" : {"templatePath" : ""}}
            return {"status" : 200,"code" : "OK" , "result" : {"templatePath" : os.path.join(STATIC_PATH, finalFileName),"templateName" : finalFileName}}

        
        return {"status" : 404,"code" : "File Error." , "result" : {"templateLinks" : ""}}
        
@app.route("/template/api/v1/validate", methods = ['POST'])
def validate():
    req_body = request.get_json()
    templateFolderPath = req_body["request"]["templatePath"]
    templateCode = req_body["request"]["templateCode"]

    # Token validation
    auth = request.headers.get("Authorization")
    signing_key = os.environ.get("SECRET_KEY")
    payload = False
    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : {"templateLinks" : ""}}
    else:
        try:
            payload = jwt.decode(auth, signing_key, algorithms=['HS256'])
        except Exception as e:
            print(e)

    if(not payload):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : {"templateLinks" : "True"}}
    

    basicErrors = xlsxObject(templateCode, templateFolderPath)

    if basicErrors.success:
        valErr = basicErrors.basicCondition()
        advValErr = basicErrors.customCondition()
        return addComments(templateFolderPath,{"status" : 200,"code" : "OK" , "result" : {"basicErrors" : valErr,"advancedErrors" : advValErr}})
    else:
        return {"status" : 404,"code" : "ERROR" , "result" :{},"message":"Please check template id"}


@app.route("/template/api/v1/errDownload", methods = ['GET'])
def errDownload():
    templateFolderPath = request.args.get("templatePath")
    return send_from_directory(os.path.dirname(templateFolderPath), os.path.basename(templateFolderPath), as_attachment=True)

@app.route("/template/api/v1/userRoles/list", methods = ['GET'])
def userRoles():
    returnResponse = {}
    subRoles = connectDb(os.environ.get('mongoURL'),os.environ.get('db'),os.environ.get('conditionsCollection'))
    returnResponseTmp = subRoles.find({"name" : "recommendedForCheck"})
    
    if returnResponseTmp:
        returnResponse["status"] = 200
        returnResponse["code"] = "OK"
        returnResponse["result"] = returnResponseTmp[0]['recommendedForCheck']['roles']
    return returnResponse

# Update and add new subroles using this API
@app.route("/template/api/v1/userRoles/update", methods = ['POST'])
def update():

    error = []
    result = {}

    req_body = request.get_json()

    auth = request.headers.get('admin-token')

    # Auth code check
    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : []}
    else:
        if not auth == os.environ.get("admin-token"):
            return {"status" : 500,"code" : "Not Authorized" , "result" : []}
    try:
        mydict = {}

        if req_body["request"]["code"] == None or req_body["request"]["title"] == None or req_body["request"]["code"] == "" or req_body["request"]["title"] == "" or req_body["request"]["_id"] == None or req_body["request"]["_id"] == "" :
            error.append("Required value missing")
        else:
            subRoles = connectDb(os.environ.get("mongoURL"),os.environ.get("db"),os.environ.get("conditionsCollection"))
            returnResponseTmp = subRoles.find({"name" : "recommendedForCheck"})
    
            mydict = {"code" : req_body["request"]["code"] , "title" : req_body["request"]["title"], "_id" : req_body["request"]["_id"]}
            chechSubRole = 0
            codeFlag = False
            idFlag = False

            currentSubRoles = returnResponseTmp[0]["recommendedForCheck"]["roles"]
            for index in currentSubRoles:
                if index["code"] == req_body["request"]["code"]:
                    codeFlag = True
                if index["_id"] == req_body["request"]["_id"]:
                    idFlag = True

            if codeFlag:
                error.append("Duplicate Code error")

            # update Subrole 
            if idFlag :
                indexCount = 0
                for index in currentSubRoles:
                    if index["_id"] == req_body["request"]["_id"]:
                        currentSubRoles[indexCount]["code"] = req_body["request"]["code"]
                        currentSubRoles[indexCount]["title"] = req_body["request"]["title"]
                    indexCount += 1
                findQuery = { "name" : "recommendedForCheck" }
                newvalues = { "$set": { "recommendedForCheck.roles": currentSubRoles } }

                subRoles.update_one(findQuery, newvalues)
                result = {
                    "message" : "subRoles updated successfully.",
                    "_id" : req_body["request"]["_id"],
                    "title" : req_body["request"]["title"],
                    "code" : req_body["request"]["code"]
                }
                # Add new subroles   
            elif len(error) <= 0:

                currentSubRoles.append(mydict)
                findQuery = { "name" : "recommendedForCheck" }
                newvalues = { "$set": { "recommendedForCheck.roles": currentSubRoles } }

                subRoles.update_one(findQuery, newvalues)
                result = {
                    "message" : "subRoles added successfully."
                }
    except Exception as e:
        print(e)
        error = "Key missing."

    
    return {"status" : 200,"code" : "OK", "result" : result,"error" : error}


# @app.route("/support/api/v1/userRoles/bulkUpload", methods = ['POST'])
# def bulkUpload():

#     error = ""

#     req_body = request.get_json()

#     req_body = req_body['content']

#     result = {
#         "data"  : []
#     }
                

#     auth = request.headers.get('admin-token')

#     if(not auth):
#         return {"status" : 500,"code" : "Authorization Failed" , "result" : []}
#     else:
#         if not auth == os.environ.get('admin-token'):
#             return {"status" : 500,"code" : "Not Authorized" , "result" : []}
#     try:
#         mydict = []

#         subRoles = connectDb(os.environ.get('mongoURL'),os.environ.get('db'),os.environ.get('subRoleCollection'))
        
#         for index in req_body:
#             chechSubRole = subRoles.count_documents({"code" : index['code']})
#             if chechSubRole > 0:
#                 error = "Duplicate Key error"
#             else:
#                 error = None
#                 if index['code'] == None or index['title'] == None or index['code'] == "" or index['title'] == "":
#                     error = "Required value missing"
#                 else:
#                     error = None
#             if not error:
#                 mydict.append({"code" : index['code'] , "title" : index['title'],"error": None})
#             else:
#                 mydict.append({"code" : index['code'] , "title" : index['title'],"error" : error})

#         for index in mydict:

#             if index["error"]:
#                 result['data'].append({
#                                     "code" : index['code'],
#                                     "title" : index['title'],
#                                     "_id" : "",
#                                     "error" : index['error']
#                                 })
#             else:
#                 x = subRoles.insert_one({"code" : index['code'] , "title" : index['title']})
#                 result['data'].append({
#                                         "code" : index['code'],
#                                         "title" : index['title'],
#                                         "_id" : str(x.inserted_id),
#                                         "error" : ""
#                                     })

#     except Exception as e:

#         print(e)
        
#         error = "Key missing."

    
#     return {"status" : 200,"code" : "OK", "result" : result,"error" : error}

if (__name__ == '__main__'):
    app.run(debug=False)