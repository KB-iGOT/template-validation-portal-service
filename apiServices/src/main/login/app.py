#write your code
from flask import Flask, request
import os
from dotenv import load_dotenv
import json
import hashlib
import jwt


app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(BASE_DIR, '.env')  # just an e.g

if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
else:
    import sys
    print('".env" is missing.')
    sys.exit(1)


@app.route("/api/v1/authenticate", methods = ['POST'])
def login():
    req_body = request.get_json()
    savedUSERNAME = os.environ.get('user')
    savedPASSWORD = os.environ.get('password')
    userName = req_body['email']
    password = hashlib.md5(req_body['password'].encode('utf-8'))

    if(userName==savedUSERNAME and password.hexdigest()==savedPASSWORD):
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

if (__name__ == '__main__'):
    app.run(debug=False)