programColumnJson = "/home/piyush/Desktop/SL code/template-validation-portal-service/backend/src/main/modules/programColumnConfig.json"
programConditonalJson = "/home/piyush/Desktop/SL code/template-validation-portal-service/backend/src/main/modules/programConditioalCoulmn.json"

connectionUrl="mongodb://localhost:27017/"
databaseName = "templateValidation"
collectionName = "validation"
conditionCollection = "conditions"

hostUrl = "https://diksha.gov.in/"
preprodHostUrl = "https://preprod.ntp.net.in/"
tokenApi = "auth/realms/sunbird/protocol/openid-connect/token"
tokenHeader = {
        "Content-Type" : "application/x-www-form-urlencoded"
}
tokenData = {
        "client_id":"samiksha-app",
        "username" : "dikshatestTN1@gmail.com",
        "password" : "ShikshaLokam@123",
        "grant_type": "password",
        "client_secret": "401e939c-6743-4247-aa0f-c6d56ff5a742"
}