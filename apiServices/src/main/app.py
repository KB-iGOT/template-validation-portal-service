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

@app.route("/template/api/v1/authenticate", methods = ['POST'])
def login():
    req_body = request.get_json()
    savedUSERNAME = os.environ.get('email')
    savedPASSWORD = os.environ.get('password')
    userName = req_body['email']
    password = hashlib.md5(req_body['password'].encode('utf-8'))

    if(userName==savedUSERNAME and password.hexdigest()==savedPASSWORD):
        # Exipry and other details can be added here
        message = {
            'iss': '',
            'email': savedUSERNAME
            }
        signing_key = os.environ.get("SECRET_KEY")
        encoded_jwt = jwt.encode({'message': message}, signing_key, algorithm='HS256')


        return {"status" : 200,"code" : "Authenticated","errorFlag" : False,"error" : [],"response" : {
            "accessToken" : encoded_jwt
        }}
    else:
        return {"status" : 404,"code" : "Error","errorFlag" : True,"error" : ["Username / Password Doesn't Match"],"response" : {
            "accessToken" : "" }}

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
    templateFolderPath = req_body['templatePath']
    templateCode = req_body['templateCode']

    basicErrors = basicValidation(templateFolderPath,templateCode)
    # advancedErrors = advancedValidation(templateFolderPath,templateCode)

    if basicErrors.success:
        valErr = basicErrors.basicCondition()
        advValErr = basicErrors.customCondition()
        return {"status" : 200,"code" : "OK" , "result" : {"basicErrors" : valErr,"advancedErrors" : advValErr}}
    else:
        return {"status" : 404,"code" : "ERROR" , "result" :{},"message":"Please check template id"}
def basicValidation(templateFolderPath,templateCode):
    return xlsxObject(templateCode, templateFolderPath)


def advancedValidation(templateFolderPath,templateCode):
    return {"errors" : ["a","b","c"]}
    

if (__name__ == '__main__'):
    app.run(debug=False)
    