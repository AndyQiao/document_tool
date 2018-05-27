import xlwt;  
import xlrd; 
from xlutils.copy import copy  


oldWb=xlrd.open_workbook('shop_president_source1.xlsx');  
newWb=copy(oldWb);  
newWs=newWb.get_sheet(0);  
newWs.write(41,17,"贾倩");  

newWb.save('shop_president_source11.xlsx');  
  