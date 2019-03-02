import pandas
import xlrd
import numpy

"""Main function"""
def main():

    """Read dummy data using file handling function"""
    students = inputStudentFileHandling()
    projects = inputProjectFileHandling()

    """Empty dictionary to hold matched students and projects"""
    matchedNames = []
    matchedTitles = []
    matchedPairs = {'Student Names':matchedNames, 'Project Titles':matchedTitles}

    """Allocation function"""
    allocationLoop(students, projects, matchedPairs, matchedNames, matchedTitles)

    """Print output"""
    outputFileHandling(matchedPairs)



    

"""Functions to handle the input data"""


def inputStudentFileHandling():

    """Handles student data"""
    
    inputStudents = pandas.read_excel('inputData.xlsx', 'Students') #read excel input data file
    
    raw_names = inputStudents['Name'].values #assign imported data to a variable
    raw_student_numbers = inputStudents['Student_ID'].values
    raw_choices = inputStudents['Choices'].values
    raw_grades = inputStudents['Grade'].values
    
    names = [] #creates empty lists to be populated with formatted data
    student_numbers = []
    choices = []
    grades = []

    """Formats raw data by converting ndarrays into lists"""
    for z in raw_names:
        names.append(z)

    for z in raw_student_numbers:
        student_numbers.append(z)
    
    for z in raw_choices:
        z = [int(y) for y in z.split(" ")]
        choices.append(z)

    for z in raw_grades:
        grades.append(z)

    """Format all imported data into lists of list"""
    students = [names, student_numbers, choices, grades]

    return students




def inputProjectFileHandling():

    """Handles project data"""
    
    inputProjects = pandas.read_excel('inputData.xlsx', 'Projects') #read excel input data file
    
    raw_titles = inputProjects['Title'].values #assign imported data to a variable
    raw_project_numbers = inputProjects['Project_ID'].values

    titles = [] #creates empty lists to be populated with formatted data
    project_numbers = []

    """Formats raw data by converting ndarrays into lists"""
    for z in raw_titles:
        titles.append(z)

    for z in raw_project_numbers:
        project_numbers.append(z)

    """Format all imported data into lists of list"""
    projects = [titles, project_numbers]

    return projects


    
    
"""Function to run allocation function"""
def allocationLoop(students, projects, matchedPairs, matchedNames, matchedTitles):
    """Function loop"""
    loop_length = len(students[1])
    z = 0
    while z < loop_length: #loop runs the two functions that handle the matching process
        highestGrade(students) #run highest grade function
        projectAssignment(students, projects, highest_grade_index, matchedPairs, matchedNames, matchedTitles) #run project assignment function
        listCleanup(students, projects, highest_grade_index) #run list cleanup function 
        z = z+1



        

"""function to return the student with the highest grade"""
def highestGrade(students):
    global highest_grade_index
    highest_grade = 0 #variable to hold the highest grade
    for x in students[3]: #loop that counts through the grades of the students
        if highest_grade < x: #when the grade the counter points at is greater that that of the holding variable
            highest_grade = x #the value the counter is pointing at is assigned to the holding variable
            highest_grade_index = students[3].index(x) #the list index is also held in a variable
    return highest_grade_index





"""Function to assign student to project"""
def projectAssignment(students, projects, highest_grade_index, matchedPairs, matchedNames, matchedTitles):
    for y in students[2][highest_grade_index]: #loop through the highest achieving student's choices
        for x in projects[1]: #looping through the list of projects 
            if y == x: #in the case of a choice-project match
                project_index = projects[1].index(x) #index variable for the matched project

                matchedNames.append(students[0][highest_grade_index]) #append list of matched students
                matchedTitles.append(projects[0][project_index]) #append list of matched projects

                projects[0].remove(projects[0][project_index]) #remove matched project title
                projects[1].remove(projects[1][project_index]) #remove matched project number
                return
    else:
        matchedNames.append(students[0][highest_grade_index])
        matchedTitles.append('Unassigned')
        return

            
                
                
"""Function to remove already matched students and projects from the dataset"""
def listCleanup(students, projects, highest_grade_index):
    students[0].remove(students[0][highest_grade_index])
    students[1].remove(students[1][highest_grade_index])
    students[2].remove(students[2][highest_grade_index])
    students[3].remove(students[3][highest_grade_index])




"""Function to handle the output data and write it to an excel file"""
def outputFileHandling(matchedPairs):
    """Convert dictionary to pandas dataframe"""
    df = pandas.DataFrame(data = matchedPairs)
    print(df)

    """Write output to excel file"""
    df.to_excel('output.xlsx', sheet_name='Matches')



"""Function to calculate the percentage of students with a high choice"""
def choicePercentage():
    highChoice = 0
    lowChoice = 0
    highChoicePercentage = 0
    count = 0
    for y in students[2][highest_grade_index]: #loop through the highest achieving student's choices
        for x in projects[1]: #looping through the list of projects 
            if y == x: #in the case of a choice-project match
                highChoice = highChoice + 1

            
