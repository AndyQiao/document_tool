Sub modify()

aaa = Array("柯涛","苏治国","杨晓光","罗思恩","邓涛","高博瀚","刘海英","张景辉","庄毅","田殊艳","张炜","李欢","孟琦","赵怡婷","张阳","张玉东","罗靖","贺亚杰","张龙忠","文静","李林","郭艳玲","夏胜美","叶翠","魏伟","乔珊珊","王晓虎","许艾湛","钟立成","邹喜","周翔","穆慕","杜冬","许累欣","易德陶","舒雷震","潘文","许强","严炜","胡志明","李航","程思杰","李韵琴","袁涛","陈晓莹","姚显梅","李剑","刘丽霞","王小腾","旷文新")
'aaa = Array("刘磊","宋思明","易庆","叶丽娟","张权富","钟铭","何崇智","谢晶","吴诗伟","李巧红","刘欢","李若楠","徐一铭","劳林高弘","王京京","陈熠","彭卫平","梁丽娇","魏青","游崇龙","叶丽环","柯春根","范其俊","张文婷","胡军","杨秀","任伽","朱玉婷","黄斌","周畅","罗静","曹樱子","施辉","周连芳","丁继华","江宣","张艮发","柳晶","王旋","苗芸","鲁秋池","韩正坤","耿永刚","周遊","周文","肖铃慧")


For i = 0 To UBound(aaa)
Cells(42, 18) = aaa(i)
Application.DisplayAlerts = False
Dim filename As String
filename = aaa(i) + "-采购总监资深采购胜任力测评综合评估报告"
'保存Excel：ActiveWorkbook.SaveAs filename:="D:\genius_tool\document_tool\fill_excel_to_pdf\2\" + filename
 ActiveWorkbook.ExportAsFixedFormat xlTypePDF, "D:\genius_tool\document_tool\fill_excel_to_pdf\2_pdf\" & filename
Next


End Sub