from openpyxl import load_workbook
import webbrowser
workbook = load_workbook(filename="YOUR FILENAME HERE .xlsx")
sheet = workbook.active

def make_search(term):
    s = "https://www.google.com/search?q="+term
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
