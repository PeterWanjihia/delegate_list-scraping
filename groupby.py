import pandas as pd

# Read the categorized and sorted data from the Excel file
df = pd.read_excel('Categorized_and_Sorted_Applications.xlsx')

# Create a dictionary to store students for each course
students_by_course = {}

# Iterate over unique courses and collect students
for course in df['Course'].unique():
    students_by_course[course] = df.loc[df['Course'] == course, ['Name', 'Adm No', 'Position']].values.tolist()

# Save the dictionary of students by course to an Excel file
with pd.ExcelWriter('Students_by_Course.xlsx') as writer:
    for i, (course, students) in enumerate(students_by_course.items()):
        sheet_name = f'{course}_Students_{i + 1}'  # Append a number to the sheet name
        students_df = pd.DataFrame(students, columns=['Name', 'Adm No', 'Position'])
        students_df.to_excel(writer, sheet_name=sheet_name, index=False)
