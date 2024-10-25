# SharePoint List Form Automation

This project was developed to automate the form-filling process in SharePoint using data from an Excel spreadsheet. Leveraging Selenium for browser control and Tkinter for a user-friendly graphical interface, the system facilitates the repeated submission of data to SharePoint, reducing manual errors and saving time.

**KEY FEATURES**

1. **Automated Form Filling**  
   The system extracts data from an Excel spreadsheet and automatically fills out a form hosted on SharePoint. Each form field is mapped to a specific column in the spreadsheet.

2. **Graphical Interface with Tkinter**  
   Users can initiate the automation process via a simple graphical interface. This interface allows users to open the browser, access the form, and start the filling process with just a few clicks.

3. **Excel Integration**  
   Data is pulled directly from an Excel file, making the process flexible and easy to update. The system supports structured sheets with 17 columns, following the provided template.

4. **Support for Parallel Execution**  
   Form data filling occurs in a separate thread, ensuring that the graphical interface remains responsive while the automation is in progress.

**TECHNOLOGIES USED**

1. **Python**: Main programming language used in the project.
2. **Selenium**: Used to control the browser and automatically fill in form fields on SharePoint.
3. **Tkinter**: Library used to create the graphical user interface.
4. **Pandas**: Used for data manipulation and reading from Excel spreadsheets.
5. **Excel**: Data source for form filling.

**HOW IT WORKS**

1. **Excel Sheet Reading**: The system loads an Excel sheet containing the required data to fill out the form. The sheet should have the following columns:
   - DATA PROVISION, CLIENT, BUSINESS UNIT, IC, DOC TYPE, DOC NUMBER, CLASSIFICATION, GROSS REVENUE, ICMS, ISS, PIS, COFINS, CPRB, NET REVENUE, REMARKS, STATUS, LINK

2. **Browser Opening**: The system automatically opens the browser (Chrome) and accesses the SharePoint page where the form is hosted.

3. **Automated Filling**: Using data from the spreadsheet, the system fills each form field and clicks the “Submit” button. After submission, the system waits for the form to reload to process the next row of data.

4. **Graphical Interface**: The interface allows users to open the browser and start the filling process quickly and easily, all with just a few clicks.

**CONCLUSION**

This project provides a practical and efficient solution for automating form-filling in SharePoint, using data from an Excel spreadsheet. With an intuitive graphical interface and integration with Selenium, the system simplifies large-volume data entry, making it faster and less error-prone.
