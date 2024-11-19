# **SharePoint List Form Automation**

This project automates the process of filling out SharePoint forms using data from an Excel spreadsheet. By combining Selenium for browser automation and Tkinter for a user-friendly interface, the system significantly reduces manual effort and errors, streamlining repetitive data entry tasks.

---

## **KEY FEATURES**

### 1. **Automated Form Filling**
   - Extracts data from an Excel spreadsheet to populate SharePoint forms.
   - Each form field is precisely mapped to a corresponding column in the spreadsheet.

### 2. **Graphical Interface with Tkinter**
   - Offers a simple and intuitive interface for users to initiate the automation process.
   - Allows quick access to browser control, form navigation, and automation execution.

### 3. **Excel Integration**
   - Reads data directly from an Excel file with a structured format of 17 predefined columns.
   - Ensures seamless data import and easy updates through an Excel template.

### 4. **Support for Parallel Execution**
   - Automates the form-filling process in a separate thread, keeping the interface responsive during operation.

---

## **TECHNOLOGIES USED**

### 1. **Python**
   - Core programming language driving the system's functionality.

### 2. **Selenium**
   - Controls the browser to access SharePoint and fill out forms automatically.

### 3. **Tkinter**
   - Provides a graphical user interface for easy interaction with the system.

### 4. **Pandas**
   - Handles data manipulation and reading from Excel files.

### 5. **Excel**
   - Serves as the primary data source for the automation process.

---

## **HOW IT WORKS**

1. **Excel Sheet Reading**  
   - Loads data from an Excel file structured with the following columns:  
     - **DATA PROVISION, CLIENT, BUSINESS UNIT, IC, DOC TYPE, DOC NUMBER, CLASSIFICATION, GROSS REVENUE, ICMS, ISS, PIS, COFINS, CPRB, NET REVENUE, REMARKS, STATUS, LINK**

2. **Browser Opening**  
   - Automatically launches Chrome and navigates to the SharePoint form page.

3. **Automated Form Filling**  
   - Populates each form field using data from the spreadsheet.  
   - Submits the form and waits for it to reload before processing the next row.

4. **Graphical Interface**  
   - Provides buttons to open the browser, start the automation, and monitor progress, offering a smooth user experience.

---

## **CONCLUSION**

The **SharePoint List Form Automation** system provides an efficient and error-free solution for bulk data entry into SharePoint. By combining a responsive graphical interface with robust browser automation, it simplifies the process of submitting large datasets, saving time and ensuring data accuracy. Ideal for organizations handling repetitive SharePoint tasks, this tool enhances productivity and minimizes manual effort.
