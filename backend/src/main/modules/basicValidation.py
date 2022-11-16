#write your code
import sys
sys.path.insert(0,"/home/piyush/Desktop/SL code/template-validation-portal-service/backend/src/main/")
from modules.xlsxObject import xlsxObject

xlsx1 = xlsxObject(id="3", xlsxPath="/home/piyush/Desktop/SL code/template-validation-portal-service/backend/src/main/VAM_CHD_ProgramTemplate.xlsx")
# print(xlsx1.checkSheetExists())
print(xlsx1.basicCondition())
# print(xlsx1.customCondition())


