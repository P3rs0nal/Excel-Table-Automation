# Excel-Automation

This code is used to generate email drafts from an excel table given a certain structure. The purpose of this code is to expedite the process of table data anaylsis. Using the first few columns of the table on Email Generator, the macro "Draft Email" will compile the table into a list of User Objects. An email is drafted from each user's information. The version of email generation is determined by a button on the "Email Generator" sheet. This button will toggle between a word, email draft, and direct email send version. The email is formatted accordingly. A table of report history can be found in the sheet "Email History", listing the times a report was generated per user. The language of the email can be customized in the sheet "Customized Language." The customized language gets sent to an HTML converter to allow for dynamic formatting of any language. The sheet "How to Use" contains instructions on each button and a general overview of "Draft Email."

# Purpose

The purpose of this project is to automate information requests that follow similar structure among multiple users.

# Further Implementation

This project is capable of being implemented into a 3 part design: Data review / collection, data request automation, and response collection automation. The code is easily transmutable in addition to Excel compatible with most softwares through add-in/plugins. 
