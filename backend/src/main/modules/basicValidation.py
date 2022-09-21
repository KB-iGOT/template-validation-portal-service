#write your code
import sys
sys.path.insert(0,"/home/piyush/Desktop/SL code/template-validation-portal-service/backend/src/main/")
from modules.xlsxObject import new_xlsxObject

xlsx1 = new_xlsxObject(id="1")
print(xlsx1.getSheetNames())
print(xlsx1.checkSheetExists())
print(xlsx1.checkColumnsExists())
print(xlsx1.checkDates())
