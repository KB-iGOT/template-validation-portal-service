#write your code
from flask import Flask, request , send_from_directory, jsonify
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
from pymongo.errors import DuplicateKeyError


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
                spreadSheet = workBook[errData["sheetName"]]
                try:
                    columnNumer = xlsxData[errData["sheetName"]].columns.get_loc(errData["columnName"])
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
                        if spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment is None:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                        else:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment=Comment(spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                elif type(errData["rowNumber"]) is int:
                    if spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment is None:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
                    else:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment=Comment(spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid") 
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
    

    basicErrors = basicValidation(templateFolderPath,templateCode)
    # advancedErrors = advancedValidation(templateFolderPath,templateCode)

    if basicErrors.success:
        valErr = basicErrors.basicCondition()
        advValErr = basicErrors.customCondition()
        return addComments(templateFolderPath,{"status" : 200,"code" : "OK" , "result" : {"basicErrors" : valErr,"advancedErrors" : advValErr}})
    else:
        return {"status" : 404,"code" : "ERROR" , "result" :{},"message":"Please check template id"}
def basicValidation(templateFolderPath,templateCode):
    return xlsxObject(templateCode, templateFolderPath)


def advancedValidation(templateFolderPath,templateCode):
    return {"errors" : ["a","b","c"]}
    


@app.route("/template/api/v1/errDownload", methods = ['GET'])
def errDownload():
    templateFolderPath = request.args.get("templatePath")
    return send_from_directory(os.path.dirname(templateFolderPath), os.path.basename(templateFolderPath), as_attachment=True)

@app.route("/support/api/v1/userRoles/list", methods = ['GET'])
def userRoles():
    returnResponse = {}
    subRoles = connectDb(os.environ.get('mongoURL'),os.environ.get('db'),os.environ.get('subRoleCollection'))
    returnResponseTmp = subRoles.find({})
    if returnResponseTmp:
        returnResponse["status"] = 200
        returnResponse["code"] = "OK"
        returnResponse["result"] = []
        for ech in returnResponseTmp:
            returnResponse["result"].append({"_id" : str(ech["_id"]),"code" : ech["code"],"title" : ech["title"]})
    return returnResponse


@app.route("/support/api/v1/userRoles/singleUpload", methods = ['POST'])
def singleUpload():

    error = ""
    result = {}

    req_body = request.get_json()

    auth = request.headers.get('admin-token')
    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : []}
    else:
        if not auth == os.environ.get('admin-token'):
            return {"status" : 500,"code" : "Not Authorized" , "result" : []}
    try:
        mydict = {}

        if req_body['code'] == None or req_body['title'] == None or req_body['code'] == "" or req_body['title'] == "":
            error = "Required value missing"
        else:
            subRoles = connectDb(os.environ.get('mongoURL'),os.environ.get('db'),os.environ.get('subRoleCollection'))
            mydict = {"code" : req_body['code'] , "title" : req_body['title']}
            chechSubRole = subRoles.count_documents({"code" : req_body['code']})

            if chechSubRole > 0:
                error = "Duplicate Key error"
            else:

                x = subRoles.insert_one(mydict)
                result = {
                    "message" : "subRoles added successfully.",
                    "_id" : str(x.inserted_id)
                }
    except Exception as e:
        
        error = "Key missing."

    
    return {"status" : 200,"code" : "OK", "result" : result,"error" : error}


@app.route("/support/api/v1/userRoles/bulkUpload", methods = ['POST'])
def bulkUpload():

    error = ""

    req_body = request.get_json()

    req_body = req_body['content']

    result = {
        "data"  : []
    }
                

    auth = request.headers.get('admin-token')

    if(not auth):
        return {"status" : 500,"code" : "Authorization Failed" , "result" : []}
    else:
        if not auth == os.environ.get('admin-token'):
            return {"status" : 500,"code" : "Not Authorized" , "result" : []}
    try:
        mydict = []

        subRoles = connectDb(os.environ.get('mongoURL'),os.environ.get('db'),os.environ.get('subRoleCollection'))
        
        for index in req_body:
            chechSubRole = subRoles.count_documents({"code" : index['code']})
            if chechSubRole > 0:
                error = "Duplicate Key error"
            else:
                error = None
                if index['code'] == None or index['title'] == None or index['code'] == "" or index['title'] == "":
                    error = "Required value missing"
                else:
                    error = None
            if not error:
                mydict.append({"code" : index['code'] , "title" : index['title'],"error": None})
            else:
                mydict.append({"code" : index['code'] , "title" : index['title'],"error" : error})

        for index in mydict:

            if index["error"]:
                result['data'].append({
                                    "code" : index['code'],
                                    "title" : index['title'],
                                    "_id" : "",
                                    "error" : index['error']
                                })
            else:
                x = subRoles.insert_one({"code" : index['code'] , "title" : index['title']})
                result['data'].append({
                                        "code" : index['code'],
                                        "title" : index['title'],
                                        "_id" : str(x.inserted_id),
                                        "error" : ""
                                    })

    except Exception as e:

        print(e)
        
        error = "Key missing."

    
    return {"status" : 200,"code" : "OK", "result" : result,"error" : error}

if (__name__ == '__main__'):
    app.run(port=8000,debug=False)