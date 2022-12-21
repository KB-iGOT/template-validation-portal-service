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

def addComments(templatePath, errResponse):
    xlsxData = pd.read_excel(templatePath, sheet_name=None)
    errPath = templatePath.split(".")[0]+"_errFile"+".xlsx"
    shutil.copyfile(templatePath, errPath)

    workBook = load_workbook(errPath, data_only = True)
    
    for key in xlsxData.keys():
        newHeader = xlsxData[key].iloc[0]
        xlsxData[key] = xlsxData[key][1:]
        xlsxData[key].columns = newHeader
    for result in errResponse["result"]:
        for errData in errResponse["result"][result]["data"]:
            if errData["columnName"] != "":
                spreadSheet = workBook[errData["sheetName"]]
                try:
                    columnNumer = xlsxData[errData["sheetName"]].columns.get_loc(errData["columnName"])
                except Exception as e:
                    if spreadSheet.cell(2,1).comment is None:
                        spreadSheet.cell(2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                    else:
                        spreadSheet.cell(2,1).comment=Comment(spreadSheet.cell(row=1, column=1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                    continue
                if type(errData["rowNumber"]) is list:
                    for rowIndex in errData["rowNumber"]:
                        if spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment is None:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                        else:
                            spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment=Comment(spreadSheet.cell(row=rowIndex+2, column=columnNumer+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")

                elif type(errData["rowNumber"]) is int:
                    if spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment is None:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
                    else:
                        spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment=Comment(spreadSheet.cell(row=errData["rowNumber"]+2, column=columnNumer+1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n","admin")
            else:
                if errData["errCode"] == 300:
                    workBook.create_sheet(errData["sheetName"])
                    spreadSheet = workBook[errData["sheetName"]]
                    spreadSheet.cell(2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                    continue
                else:
                    spreadSheet = workBook[errData["sheetName"]]
                    for rowIndex in errData["rowNumber"]:
                        if spreadSheet.cell(rowIndex+2,1).comment is None: 
                            spreadSheet.cell(rowIndex+2,1).comment=Comment("Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
                        else:
                            spreadSheet.cell(rowIndex+2,1).comment=Comment(spreadSheet.cell(rowIndex+2,1).comment.text+"Error - "+errData["errMessage"]+"\n Suggestion -"+errData["suggestion"]+"\n" ,"admin")
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

if (__name__ == '__main__'):
    app.run(debug=False)
