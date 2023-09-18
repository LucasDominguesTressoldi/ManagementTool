# CustomersManagementTool 🗃️

**An application that I developed while working freelance for a tutoring school. The application is capable of managing students and teachers, as well as registering them, checking attendance in classes, etc.**

## 🎯 What is the purpose of the application?

This application was developed in Python, using some libraries such as:  
openpyxl - customtkinter - CTkMessageBox - sys - os - datetime - among others;

Its function is simple: Generate an interface that allows the user to edit specific Excel files in a simple and practical way.

The application aims, above all, to remove the repetitiveness of creating new tables every month for customers and employees -- in this case, students and teachers.

*It won't be difficult, if anyone wants, to modify the code for another business niche other than a tutoring school.*

## ⚙️ How it works?

Inside the "assets" folder are all the files that the application uses to function, among these files are:  
the backup folder, the contracts folder, Excel files, .json files for customizing a customtkinter theme, and an image in. ico to be the interface and .exe icon (will be commented below).

The application contains 4 buttons named:  
New Student, Check Attendance, New Teacher, and Delete Student/Teacher.

***New Student*** 🆕

This button is responsible for creating a new student.

A screen will display the student's data and define their classes (day, time, teacher, subject, price, etc.). Then, the user has the option to choose what they want to do:
Create a contract for the student, Generate a spreadsheet for the student, Register it in the student's Excel file, or all simultaneously.

If the option to Generate a spreadsheet for the student is selected, the data will also be automatically inserted into the teachers' spreadsheets. In other words, if a student has a class on day X with teacher Y, this information will automatically be 
inserted into the teacher's spreadsheet.

***New Teacher*** 🆕

This button just creates a spreadsheet for the new teacher. Data will be added as students for that teacher are registered.

***Check Attendance*** ✅

*As it is an application for a tutoring school, it is very important that you can view the presence of students and teachers in classes.*

Therefore, this button requests student data such as class day, time, and subject and allows the user to select between three options: *Done, Pending, and Missed (if the user wants, they can also enter any other text to be added in place of the defaults).*

Then, the application looks for data that matches the data provided and in the **Attendance** column changes its value from empty to the one selected by the user. This is also done in the teacher's spreadsheet, but automatically.

***Delete Teacher/Student*** ❌

This button only removes the worksheet from the teacher or student provided by the user.

### ➕ Other functions

A function that is performed automatically by the application is to perform **monthly backups**.  
The application checks every time it is opened whether the year, month, and/or day correspond to the last backup performed. If one of the options is not matched, the backup is performed. In other words, backups are daily and overwrite the old ones (in the case of days). If the year and/or month are different, a new folder will be created with the name of the year and/or month.

After that, if the month has been backed up (a new monthly folder was created) the application will also understand that **it needs to create new tables for students and teachers**, recalculating class dates and resetting last month's values.

### 📁 Create an exe file

If you want to create an executable file for the application, make sure you have a recent version of Python 3 installed on your machine, fork that repository and follow the steps below:

pip install customtkinter  
pip install ctkmessagebox  
pip install pyinstaller  

To find out where customtkinter and CTkMessageBox were installed on your machine, run:  
*pip show customtkinter* and *pip show ctkmessagebox*

To create the executable run:  
pyinstaller --noconfirm --onedir --windowed --add-data "<your-path-to-customtkinter>/customtkinter;customtkinter/" --add-data "<your-path-to-ctkmessagebox>/CTkMessagebox;CTkMessagebox/" --add-data "<your-path-to-assetsFolder>/assets;assets/" --icon "assets\app-icon.ico" "app.py"

🎉 **That is all!  
😁 Thanks for reading.**
