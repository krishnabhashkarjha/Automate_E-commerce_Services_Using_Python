#
#
#
#
# def get(adaptor,doc):
#     Quickbooks={"810":["SPS QuickBooks Adaptor | RSX 7.2 | 810 - Legacy","106968"],"850":["SPS QuickBooks Adaptor | RSX 7.2 | 850 - Legacy","106967"],"875":["SPS QuickBooks Adaptor | RSX 7.2 | 875 - Legacy","110240"],"856":["SPS Quickbooks Adapter RSX 7.2 | OzLink | 856","135124"]}
#     Fishbowl={"810":["SPS Fishbowl Adaptor | RSX 7.2 | 810 - Legacy","109097"],"850":["SPS Fishbowl Adaptor | RSX 7.2 | 850 - Legacy","109095"],"875":["SPS FISHBOWL ADAPTOR | RSX 7.2 | 875 - LEGACY","112556"],"856":["SPS Fishbowl Adaptor | RSX 7.2 | 856 - Legacy","109096"]}
#     Dwyer={"810":["Dwyer Adaptor V7 810 XML","31788"],"850":["Dwyer Adaptor V7 850 XML","31768"],"875":["Dwyer Adaptor V7 875 XML","70230"]}
#     Peachtree={"810":["Peachtree 810 XML","73633"],"850":["Peachtree 850 XML","73632"],"875":["Peachtree 875 XML","78248"]}
#     try:
#         if adaptor=="Quickbooks":
#             arr=Quickbooks.get(doc)
#             return arr
#         if adaptor=="Fishbowl":
#             arr=Fishbowl.get(doc)
#             return arr
#         if adaptor=="Dwyer":
#             arr=Dwyer.get(doc)
#             return arr
#         if adaptor=="Peachtree":
#             arr=Peachtree.get(doc)
#             return arr
#     except:
#         print("in excep")
#         return ['none','nnn']
# cap=get("Quickbooks","850")
# print(cap[0])

# s = "10010"
# c = float(s)
# print("After converting to integer base 2 : ", end="")
# print(c)
##############################################################
# class persone:
#     name = "bhaskii"
#     def __init__(self,age1,age):
#         self.age1 = age1
#         self.age = age
#
#     def avg(self):
#         return (self.age+self.age1)/2
#
#     @classmethod
#     def student_name(cls,name):
#         return name
#
#     @staticmethod
#     def info():
#         print("hey bhaskii")
#
#
#
# s1 = persone(21,23)
# print(s1.avg())
# print(persone.student_name("bhaskii"))
# persone.info()
########################################################
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
df = pd.read_excel(r"C:\Users\Krishnabhashkar.Jha\Desktop\weather.xlsx")
# print(df)
# a=df['A']
# print(a)
# s=pd.Series(np.random.rand())
# print(df.head(1))
# print(df.sum)
# c=df['A']
# plt.plot(b,c)
# plt.xlabel('x')
# plt.ylabel('y')
# plt.show()
# plt.close()
list = {'Name':pd.Series(['priya','sunil','pushpalata','bhaskii']),'Age':['25','58','49','24'],'Gender':['F','M','F','M'],'income':['36k','80k','0k','22k']}
data=pd.DataFrame(list)
print(data.describe())
# new=df.replace(to_replace="41.00000",value="AA")
# print(new)
# import docx2pdf
# a=docx2pdf.convert(r"C:\Users\Krishnabhashkar.Jha\Desktop\AAA.docx",r"C:\Users\Krishnabhashkar.Jha\Desktop\AAA.pdf")
# print(a)

# import numpy as np
# data = np.array(['a','b','c','d'])
# print(pd.Series(data=11,index=[1]))