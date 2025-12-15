Production Documents Management Application

Description

Android application developed in Kotlin as part of a company internship.
The goal of the application is to simplify the management of production documents related to industrial projects stored on Microsoft SharePoint (OneDrive).
The application automates project folder search, document generation, and the insertion of standard files in order to reduce manual operations and limit errors.

Features

 - Microsoft authentication using Azure AD (MSAL)
 - Project search via barcode scanning or manual input
 - Access to SharePoint folders through Microsoft Graph API
 - Automatic generation of PDF documents
 - File management (copy, rename, and organization)
 - Selection of a standard document when a required file is missing
 - Dropdown list with search functionality for standard document selection

Technologies Used

 - Kotlin
 - Android SDK
 - Android Studio
 - Microsoft Graph API
 - MSAL (Microsoft Authentication Library)
 - OkHttp
 - JSON
 - SharePoint / OneDrive

Project Structure

 - MainActivity
    - User authentication
    - Project search and selection
    - Navigation to the detail screen
 - DetailActivity
    - Document display and management
    - Verification of required files
    - Insertion and renaming of standard documents

Installation

 - Clone the repository
 - Open the project with Android Studio
 - Configure Microsoft access (Azure AD / MSAL)
 - Run the application on an emulator or an Android device

Context

This project was developed in a real industrial environment during an internship.
The application evolved based on feedback from multiple departments to match existing workflows and constraints.
