from openpyxl import load_workbook
import webbrowser
workbook = load_workbook(filename="YOUR FILENAME HERE .xlsx")
sheet = workbook.active

def make_search(term):
    s = "https://www.google.com/search?q="+term+"&sxsrf=AOaemvJvhzbx5Thy2rW00q3qiD4tFIORzA%3A1631560943795&source=hp&ei=76Q_Yc2HLsbVrtoP3J6hyAE&iflsig=ALs-wAMAAAAAYT-y__6YKdjWByfSl5qtNTMAmPU2qb-V&oq=hi&gs_lcp=Cgdnd3Mtd2l6EAMyBAgjECcyBQgAEJECMgQIABBDMgUIABCRAjIHCC4QsQMQQzIECC4QQzIECAAQQzIECAAQQzIECC4QQzINCC4QsQMQxwEQ0QMQQzoHCCMQ6gIQJzoOCC4QgAQQsQMQxwEQowI6CwgAEIAEELEDEIMBOggIABCABBCxA1C9KlieLGCnL2gBcAB4AIABaogB0wGSAQMwLjKYAQCgAQGwAQo&sclient=gws-wiz&ved=0ahUKEwiNu-2T1vzyAhXGqksFHVxPCBkQ4dUDCAk&uact=5"
    webbrowser.open(s)
def column_search(head,i,column):
    cell1 = head+str(i)
    val1 = str(sheet[cell1].value)
    while(val1!= "None"):
        cell1 = head+str(i)
        val1 = str(sheet[cell1].value)
        cell2 = column+str(i)
        val2 = str(sheet[cell2].value)
        if(val2 != "None"):
            print(val1 + ": "+val2)
        else:
            title = column+i
            title = str(sheet[title].value)
            search_term = val1 + " " +title
            make_search(search_term)
        i+=1