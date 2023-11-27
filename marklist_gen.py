import os
from imports.get_marks import get_mark
import xlsxwriter
import os.path

def main():
    """
    Generates Excel files for each student with their respective marks and other details.

    Example Usage:
    ```python
    main()
    ```

    Inputs:
    - noOfStudents: an integer representing the number of students
    - grade: a string representing the class
    - MAXMARKS: an integer representing the total marks
    - noOfsubjects: an integer representing the number of subjects
    - subjects: a list of strings representing the names of the subjects
    - ExamName: a string representing the name of the exam
    - name: a string representing the name of a student

    Outputs:
    - Excel files named 'out/{name}.xlsx' for each student, containing their details, exam details, subject names, and marks.
    """
    merge_format = workbook.add_format({
        'bold':     True,
        'align':    'center',
        'valign':   'vcenter',})

    color_merge = workbook.add_format({
        'bold':     True,
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': '#E57283',})

    blue_merge = workbook.add_format({
        'fg_color': '#bdd7ee',
        'valign':   'vcenter',})

    green_merge = workbook.add_format({
                'fg_color': 'c6e0b4',
                'align':    'center',
                'valign':   'vcenter',
            })
    yellow_merge = workbook.add_format({
                'fg_color': '#fff2cc',
                'align': 'center',
                'valign': 'center',
                'font_color':'#203764'
            })
    
    
    try:
        #taking information from the user
        noOfStudents = int(input("Enter the number of students :"))
        grade=input("Enter the Class :")
        MAXMARKS=int(input("Enter total marks:"))
        noOfsubjects=int(input("Enter the number of subjects:"))
        ExamName = input("Enter the exam name:")
        print("Enter the subject names in order")
        
        #variables for holding the data
        subjects=[]
        marks=[]

        #create the output folder
        if not os.path.isdir('./out'):
            os.makedirs('./out', exist_ok=True)
        
        #get the list of subjects
        for i in range(0,noOfsubjects):
            subjects.append(input("Enter the subject number ",str(i)+":"))
        
        #create the excel files
        for i in range(0,noOfStudents):
            name = input("name of student"+str(i)+":")
            workbook = xlsxwriter.Workbook('out/'+name+'.xlsx')       
            ws = workbook.add_worksheet()
            for sub in subjects:
                marks.append(get_mark(name,sub))    
            
            #merging the columns and adding the relevant data to the workbook
            ws.merge_range('B1:C1','',merge_format)
            ws.merge_range('B2:C2','MAHARISHI VIDYA MANDIR SR. SEC. SCHOOL',color_merge)
            ws.merge_range('B3:C3','Ingur, Erode - 52',color_merge)
            ws.merge_range('B4:C4','',color_merge)
            ws.merge_range('B6:C6',ExamName,yellow_merge)
            ws.merge_range('B8:C8','NAME:'+name,blue_merge)
            ws.merge_range('B9:C9','GRADE:'+ grade,blue_merge)
            ws.merge_range('B11:B12','SUBJECT',green_merge)
            ws.write('C11','(18.12.2020 to 24.12.2020)',green_merge)
            ws.write('C12','Max Marks:'+str(MAXMARKS),green_merge)
            val = 13
            var = 0
            for sub in subjects:
                ws.write('B'+str(val),sub,yellow_merge)
                ws.write('C'+str(val),marks[var],workbook.add_format({
                    'align':'center'
                }))
                val+=1
                var+=1
            
            ws.write('B'+str(val+1),'Total',yellow_merge)
            ws.write('C'+str(val+1),sum(marks))
            ws.write('B'+str(val + 2),'Percentage',yellow_merge)
            ws.write('C'+str(val +2),round(sum(marks)/(MAXMARKS*noOfsubjects)*100,2))
            workbook.close()   
    except KeyboardInterrupt:
        print("\nClosing")
main()