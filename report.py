import pandas as pd
import numpy as np
import fpdf as pdf
import statistics
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
import ctypes


MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Welcome to Wisdom Tests Report Generation! Please click OK to start report generation', 'WISDOM TESTS', 0)
#### opening excel sheet using sheet name and filename and openpyxl engine, because xldr is deprecated and older engine

my_sheet = 'Sheet1' # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
file_name = 'Dummy Data.xlsx' # change it to the name of your excel file
df1 = pd.read_excel(file_name,my_sheet,engine="openpyxl",header=0)
head = 1

#### While loop to skip initial blank lines for data consistency and having a flexible system
while(df1.columns[0]=="Unnamed: 0"):
    df1 = pd.read_excel(file_name,my_sheet,engine="openpyxl",header=1)
    head+=1

#### renaming first column header for easier data extraction
df1 = df1.rename(columns={"Candidate No. (Need not appear on the scorecard)":"candidate_num"})

#### Removing extra whitespaces from the header names 
for k in df1.columns:
    df1 = df1.rename(columns={k:k.strip(' ')})

total_rows = len(df1.index)
unique_students = len(set(df1['candidate_num']))
total_questions = int(total_rows/unique_students)

total_class_marks = []
city_dict = {}
country_dict = {}

final_info = []
for j in range(0,unique_students):
    total_marks = 0
    marks_scored = 0
    correct_ans = 0
    incorrect = []
    correct = []
    unattempted_no = 0

    for i in range(j*total_questions,j*total_questions+total_questions):

        if df1["City of Residence"][i] not in city_dict.keys():
            city_dict[df1["City of Residence"][i]] = []
        if df1["Country of Residence"][i] not in country_dict.keys():
            country_dict[df1["Country of Residence"][i]] = []

        total_marks += df1['Score if correct'][i]
        if(df1['Outcome (Correct/Incorrect/Not Attempted)'][i]=="Unattempted"):
            unattempted_no+=1
        if (df1['What you marked'][i]==df1['Correct Answer'][i]):
            correct_ans+=1
            correct.append(int(df1['Question No.'][i][1:]))
        else:
            incorrect.append(int(df1['Question No.'][i][1:]))
        marks_scored += df1['Your score'][i]
    total_class_marks.append(marks_scored)
    final_info_individual = {}
    final_info_individual["Name"]=str(df1['Full Name'][i])
    final_info_individual["Registration Number"]=str(df1['Registration Number'][i])
    final_info_individual["Grade"]=str(df1["Grade"][i])
    final_info_individual["Name of School"]=str(df1['Name of School'][i])
    final_info_individual["City"] = str(df1["City of Residence"][i])
    final_info_individual["Country"] = str(df1["Country of Residence"][i])
    final_info_individual["Residence"]=str(df1["City of Residence"][i]+", "+df1["Country of Residence"][i])
    final_info_individual["Date of Birth"]=str(df1["Date of Birth"][i])
    final_info_individual["Gender"]=str(df1["Gender"][i])
    final_info_individual["Date and Time of Test"]=str(df1["Date and time of test"][i])
    final_info_individual["Total Questions"]=str(total_questions)
    final_info_individual["Total Questions attempted"]=str(total_questions-unattempted_no)
    final_info_individual["Total Questions unattempted"] = str(unattempted_no)
    final_info_individual["Total Correct attempts"]=str(correct_ans)
    final_info_individual["Total marks scored"]=str(marks_scored)
    final_info_individual["Total marks"]=str(total_marks)
    final_info_individual["Accuracy"]=str(round(100*correct_ans/(total_questions-unattempted_no),2))+'%'
    final_info_individual["Percentage"]=str(100*(marks_scored/total_marks))
    final_info_individual["Percentile"]=str(0)
    final_info_individual["Rank"]=str(0)
    final_info_individual["Average Marks"]=str(0)
    final_info_individual["Class Median"]=str(0)
    final_info_individual["Incorrect Questions"]=incorrect
    final_info_individual["Correct Questions"]=correct
    final_info_individual["City Rank"]=str(0)
    final_info_individual["Country Rank"]=str(0)
    final_info_individual["Final Result"] = str(df1["Final result"][i])
    final_info.append(final_info_individual)
    ########################################################

final_correct = [0 for i in range(int(final_info_individual["Total Questions"]))]
final_incorrect = [0 for i in range(int(final_info_individual["Total Questions"]))]

avg_marks = round((sum(total_class_marks)/len(total_class_marks)),2)
median_marks = round(statistics.median(total_class_marks),2)

def create_analytics_report(k):

    #### CREATE A PDF OBJECT
    pdf = FPDF() # A4 (210 by 297 mm)
    plt.clf()

    #### PLOTTING PIE CHART
    y = np.array([float(k["Total Correct attempts"]), float(k["Total Questions unattempted"]),float(len(k["Incorrect Questions"]))])
    mylabels = ["Correct-"+str(k["Total Correct attempts"]), "Unattempted-"+str(k["Total Questions unattempted"]), "Incorrect-"+str(len(k["Incorrect Questions"]))]
    myexplode = [0.1, 0, 0]
    plt.title("Question Distribution")
    plt.pie(y, labels = mylabels, startangle = 90,explode = myexplode, shadow = True)
    plt.savefig("./reports/piecharts/"+str(k["Registration Number"])+".png", bbox_inches='tight')
    plt.clf()

    #### PLOTTING BAR GRAPH
    x = np.array(["Class Average", "Your Marks", "Class median"])
    y = np.array([float(k["Average Marks"]),float(k["Total marks scored"]),float(k["Class Median"])])
    mycolors=["#1f50cc","#4dab2e","#ba1496"]
    plt.title("Marks Distribution")
    plt.ylabel('Marks')
    plt.bar(x,y,color=mycolors,width=0.7)
    plt.savefig("./reports/bargraphs/"+str(k["Registration Number"])+".png", bbox_inches='tight')

    plt.clf()
    #### HISTOGRAM

    width = 0.6
    x = np.arange(len(final_correct))
    p1 = plt.bar(x, final_correct, width, color='g')
    p2 = plt.bar(x, final_incorrect, width, color='r', bottom=final_correct)
    plt.ylabel('No. of people')
    plt.xlabel('Questions')
    q_list = ["Q."+str(i+1) for i in range(len(final_correct))]
    plt.xticks(x, q_list)
    plt.xticks(rotation=90)
    plt.yticks(np.arange(0,max(final_correct)+max(final_incorrect) , 1))
    plt.title('Question-wise analysis of test')
    plt.legend((p2[0], p1[0]), ('People who got it Incorrect', 'People who got it Correct'))
    plt.savefig("./reports/doublebar/"+str(k["Registration Number"])+".png", bbox_inches='tight')
    plt.clf()

    #### FIRST PAGE OF REPORT
    ''' First Page '''
    pdf.add_page()
    pdf.image("./Report-Design/header.png", 0, -5,w=210,link='http://wisdomtests.com/')
    pdf.image("./Pics for assignment/"+str(k["Name"])+".png", 133, 58,h=47)
    pdf.set_font("Courier","",15)
    pdf.set_y(60)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Name: '+str(k["Name"]), border = 0, ln = 1,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Register No.: '+str(k["Registration Number"]), border = 0, ln = 2,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Grade: '+str(k["Grade"]), border = 0, ln = 3,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Test Date: '+str(k["Date and Time of Test"]), border = 0, ln = 4,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Gender: '+str(k["Gender"]), border = 0, ln = 5,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'Residence: '+str(k["Residence"]), border = 0, ln = 6,align = 'l', fill = False)
    pdf.set_x(20)
    pdf.cell(w=20, h = 10, txt = 'D.O.B: '+str(k["Date of Birth"]),ln=7 ,border = 0, align = 'l', fill = False)
    pdf.cell(w=20, h = 10, txt = 'School: '+str(k["Name of School"]), ln=8,border = 0,align = 'l', fill = False)
    pdf.image("./Report-Design/line.png", 0, 144,w=210)
    pdf.image("./Report-Design/table1.png", 20, 150,w=170)
    pdf.set_y(172)
    pdf.set_x(155)
    pdf.cell(w=10, h = 13, txt = str(k["Total Questions"]),fill=False,border = 0,align = 'l')

    pdf.set_y(186)
    pdf.set_x(155)
    pdf.cell(w=10, h = 13, txt = str(k["Total Questions attempted"]),fill=False,border = 0,align = 'l')

    pdf.set_y(200)
    pdf.set_x(155)
    pdf.cell(w=10, h = 13, txt = str(k["Total Correct attempts"]),fill=False,border = 0,align = 'l')

    pdf.set_y(213.5)
    pdf.set_x(155)
    pdf.cell(w=10, h = 13, txt = str(int(k["Total Questions attempted"])-int(k["Total Correct attempts"])),fill=False,border = 0,align = 'l')

    pdf.set_y(228)
    pdf.set_x(151)
    pdf.cell(w=10, h = 13, txt = str(k["Total marks scored"]+'/'+k["Total marks"]),fill=False,border = 0,align = 'l')

    pdf.set_y(242)
    pdf.set_x(155)

    for i in country_dict[k["Country"]]:
        if i["Name"]==k["Name"]:
            pdf.cell(w=10, h = 13, txt = str(i["Rank"]),fill=False,border = 0,align = 'l')

    pdf.set_y(255.5)
    pdf.set_x(155)
    for j in city_dict[k["City"]]:
        if j["Name"]==k["Name"]:
            pdf.cell(w=10, h = 13, txt = str(j["Rank"]),fill=False,border = 0,align = 'l')

    pdf.set_y(270)
    pdf.set_x(151.5)
    pdf.cell(w=10, h = 7, txt = str(k["Accuracy"]),fill=False,border = 0,align = 'l')

    ''' SECOND PAGE '''
    pdf.add_page()
    pdf.image("./Report-Design/header2.png", 0, -5,w=210,link='http://wisdomtests.com/')
    pdf.image("./Report-Design/table2.png", 19, 37,w=170)

    pdf.set_y(41)
    pdf.set_x(147)
    pdf.cell(w=10, h = 10, txt = str(k["Average Marks"])+'/'+k["Total marks"],fill=False,border = 0,align = 'l')

    pdf.set_y(55)
    pdf.set_x(147)
    pdf.cell(w=10, h = 10, txt = str(k["Class Median"])+'/'+k["Total marks"],fill=False,border = 0,align = 'l')

    pdf.set_y(69)
    pdf.set_x(153)
    pdf.cell(w=10, h = 10, txt = str(k["Percentage"]+'%'),fill=False,border = 0,align = 'l')

    #### ADDING PIECHART
    pdf.image("./reports/piecharts/"+str(k["Registration Number"]+".png"), 46, 84,w=118)

    #### ADDING BARPLOT
    pdf.image("./reports/bargraphs/"+str(k["Registration Number"]+".png"), 33, 179,w=144)

    #### THIRD PAGE
    #### ADDING DOUBLE BAR GRAPH
    pdf.add_page()
    pdf.image("./reports/doublebar/"+str(k["Registration Number"]+".png"), 10, 15,w=180)
    #### ADDING FINAL
    pdf.image("./Report-Design/final.png",0, 173,w=210)
    pdf.set_y(205)
    pdf.set_x(22)
    pdf.multi_cell(w=165, h = 7.5, txt = k["Final Result"],fill=False,border = 0,align = 'l')
    #### ADDING FOOTER
    pdf.image("./Report-Design/line.png", 0, 238,w=210)
    pdf.image("./Report-Design/footer.png", 0, 255,w=210,link="http://wisdomtests.com/contact-us.html")

    pdf.add_page()
    pdf.image("./Report-Design/last.png", 0, 0,w=210,link="http://wisdomtests.com")
    #### Saving PDF by reg. nos. of students
    pdf.output("reports/"+str(k["Registration Number"])+".pdf")





#### Finding city and country rank
for k in final_info:
    k["Average Marks"] = str(avg_marks)
    k["Class Median"] = str(median_marks)
    temp = {"Name":k["Name"],"Marks":k["Total marks scored"],"Rank":0}
    temp1 = {"Name":k["Name"],"Marks":k["Total marks scored"],"Rank":0}
    city_dict[k["City"]].append(temp)
    country_dict[k["Country"]].append(temp1)

for k in city_dict:
   city_dict[k] = sorted(city_dict[k], key=lambda k: k['Marks'],reverse=True)

for k in city_dict:
    city_temp = city_dict[k]
    i=1
    for j in city_temp:
        j["Rank"]=i
        i+=1

for k in country_dict:
   country_dict[k] = sorted(country_dict[k], key=lambda k: k['Marks'],reverse=True)

i=1
for k in country_dict:
    country_temp = country_dict[k]
    i=1
    for j in country_temp:
        j["Rank"]=i
        i+=1



for k in final_info:
    for i in k["Correct Questions"]:
        final_correct[i-1]+=1
    
    for i in k["Incorrect Questions"]:
        final_incorrect[i-1]+=1

#### CREATING REPORT PDF
for k in final_info:
    create_analytics_report(k)

    
MessageBox = ctypes.windll.user32.MessageBoxW
MessageBox(None, 'Reports generated and stored in Reports Folder', 'Reports Generated', 0)