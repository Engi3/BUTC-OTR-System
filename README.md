Role: You are an Expert Full-Stack Developer specializing in Google Apps Script (GAS), Google Sheets API, HTML, CSS (Bootstrap 5), and Vanilla JavaScript.

Task: Develop the "BUTC Online Teaching Report System" for "วิทยาลัยเทคนิคบูรพาปราจีน" deployed as a Google Apps Script Web App.

Branding & UI Constraints:

    System Name: "BUTC Online Teaching Report System"

    College Name: "วิทยาลัยเทคนิคบูรพาปราจีน"

    Footer must strictly contain: "(Copyright © 2026 Assawin Namsert. All Rights Reserved.)"

    Use Bootstrap 5 for a clean, responsive, and professional UI.

Pre-defined Data (Departments):
Use this exact array of departments for all dropdowns and checkboxes:
["ช่างยนต์", "ช่างไฟฟ้า", "ช่างซ่อมบำรุง", "เมคคาทรอนิกส์", "การบัญชี", "เทคโนโลยีธุรกิจดิจิทัล", "การจัดการสำนักงานดิจิทัล", "เทคนิคพื้นฐาน"]

Backend Requirements (Code.gs - CRITICAL AUTO-INIT):

    Auto-Initialization: Write a setupSystem() function (triggered optionally or checked during doGet). It MUST automatically create the "Users" and "Reports" sheets with the exact headers if they do not exist in the active spreadsheet. It should also create a specific Google Drive folder for images if it doesn't exist.

    doGet(e) to serve the SPA Index.html.

    User Functions: registerUser() (defaults to 'Pending' status), loginUser().

    Report Functions: submitReport() (decode Base64 images, save to Drive, write to "Reports" sheet).

    Admin Functions: getPendingUsers(), approveUser(), getAllReports(), approveReport().

Frontend Requirements (Single Page Application via GAS HTML Service):

    Authentication UI: Login and Registration forms. Registration requires Email, Password, Full Name, Department (Dropdown), Phone.

    User Dashboard (Teacher):

        Report Form: >      - Teaching Date (Input date, JS MUST convert to Thai Buddhist Era +543).

            Teaching Level (Input text).

            Target Departments (Multiple Checkboxes).

            Period Selector (CRITICAL LOGIC): Visual grid for periods 1-8. P1(08:00-09:00), P2(09:00-10:00), P3(10:00-11:00), P4(11:00-12:00). Disabled "Lunch Break (12:00-13:00)". P5(13:00-14:00), P6(14:00-15:00), P7(15:00-16:00), P8(16:00-17:00). Auto-calculate total teaching hours skipping lunch.

            Photo Upload: Require 2-3 images. Convert to Base64 via FileReader.

        History Table: View submitted reports and status.

    Admin Dashboard:

        User Management Tab: List 'Pending' users with an "Approve" button.

        Reports Tab: Data table of all reports with filters (Date, Department). Include an "Approve" button and a View Image modal.

Output format: Output the complete, production-ready code blocks for Code.gs and Index.html (including inline JS/CSS). Implement loading spinners for all google.script.run calls. Ensure robust error handling.

Use Thai to communicate with me.