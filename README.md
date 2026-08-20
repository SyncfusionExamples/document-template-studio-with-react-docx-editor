# Documents Template Studio

## Introduction

Documents Template Studio is a React functional sample that provides a
centralized workspace for creating, managing, and editing DOCX document
templates using the Syncfusion<sup style="font-size:70%">&reg;</sup> [React DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/react-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) (Document Editor).

The sample is designed for donor-management and donation-processing
scenarios where organizations need reusable templates for donor
communications, receipts, reminders, and impact letters.

Users can:

-   View available document templates from a dashboard.
-   Open a template in the Syncfusion Word-like Document Editor.
-   View merge fields associated with the selected template.
-   Insert merge fields into the document.
-   Create a new blank template.
-   Upload an existing DOCX template.
-   Customize and save templates.
-   Preview a template using Mail Merge JSON data.
-   Remove templates from the application.

## Key Features

### DOCX Template Editing

Selecting a template opens it directly in the Syncfusion Document
Editor.

Users can customize the document using standard Word-like editing
capabilities, including text formatting, tables, images, and other
supported DOCX content.

### Merge Fields

The right-side merge-field panel displays the fields available for the
selected template.

Users can select a field to insert it into the document.

The sample also supports adding custom merge fields either to the
current template or to the common field catalog.

### Create and Upload Templates

Users can:

-   Create a new blank template.
-   Upload an existing `.docx` template.
-   Provide template metadata such as name, category, and purpose.
-   Open and customize uploaded templates in the editor.

### Save and Publish

The **Save and Publish** action saves the current document as a DOCX
template through the ASP.NET Core server.

The saved template is stored under:

``` text
Server-side/wwwroot/Templates/
```

### Mail Merge Preview

The **Preview with Mail Merge** action sends the current DOCX content
and JSON merge data to the server.

The ASP.NET Core server uses Syncfusion DocIO to execute the mail merge
and converts the resulting DOCX back to SFDT for display in the Document
Editor.
## Prerequisites

### Client

-   Node.js
-   npm

### Server

-   .NET 10 SDK
-   ASP.NET Core runtime
-   Syncfusion ASP.NET Core and DocIO packages referenced by the project

## How to Run

The sample has two applications:

1.  React/Vite client application
2.  ASP.NET Core server application

Start the server first because the React application uses the server for
DOCX import, save, and Mail Merge operations.

### 1. Start the ASP.NET Core Server

Open a terminal in:

``` text
Server-side/
```

Run:

``` bash
dotnet restore
dotnet run
```

The configured development URL is:

``` text
http://localhost:5212
```

The Document Editor service URL used by the React application is:

``` text
http://localhost:5212/api/DocumentEditor/
```

### 2. Start the React Application

Open another terminal in:

``` text
Client-side/
```

Install dependencies:

``` bash
npm install
```

Start the development server:

``` bash
npm run dev
```

Open the URL shown by Vite in the terminal, normally:

``` text
http://localhost:5173
```

## Template Workflow

The main workflow is:

``` text
Template Studio Dashboard
          │
          ├── Create New Template
          │
          ├── Upload DOCX Template
          │
          └── Select Existing Template
                    │
                    ▼
             Template Viewer
                    │
          ┌─────────┴─────────┐
          │                   │
          ▼                   ▼
   Document Editor       Merge Fields Panel
          │                   │
          │            Insert Merge Fields
          │                   │
          └─────────┬─────────┘
                    ▼
             Save & Publish
                    │
                    ▼
              DOCX Template
                    │
                    ▼
             Mail Merge Preview
                    │
                    ▼
              Generated SFDT
```

## Mail Merge Data

The sample uses a JSON string for Mail Merge data.

For example, a Pledge Payment Reminder can use:

``` json
{
  "Organization": [
    {
      "OrgName": "ABC Foundation",
      "OrgAddress": "123 Main Street, New York, NY 10001",
      "ReminderDate": "August 20, 2026",
      "DonorName": "John Smith",
      "DonorAddress": "45 Oak Street, New York, NY 10002",
      "PledgeNumber": "PLG-2026-0042",
      "PledgeDate": "July 15, 2026",
      "PledgeAmount": "$2,500.00",
      "AmountPaid": "$1,500.00",
      "OutstandingAmount": "$1,000.00",
      "DueDate": "September 15, 2026",
      "PaymentMethod": "Online Payment",
      "ContactEmail": "donations@abcfoundation.org",
      "ContactPhone": "+1-212-555-0123"
    }
  ]
}
```

The server receives this JSON through the `mailMergeData` property and
prepares it for Syncfusion DocIO Mail Merge processing.

## Server API

The main Document Editor endpoints used by the sample are:

| Endpoint | Purpose |
| --- | --- |
| `POST /api/DocumentEditor/Import` | Imports a DOCX and converts it to SFDT. |
| `POST /api/DocumentEditor/Save` | Saves the edited SFDT content as a DOCX template. |
| `POST /api/DocumentEditor/MailMerge` | Executes Mail Merge using the supplied JSON data and returns SFDT. |

## Resources

- **Product page:**   [Syncfusion® React DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/react-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

- **Documentation:**   [Syncfusion® React DOCX Editor - Documentation](https://help.syncfusion.com/document-processing/word/word-processor/react/overview?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

- **Online demo:**   [Syncfusion® React DOCX Editor - Online demo](https://document.syncfusion.com/demos/docx-editor/react/#/tailwind3/document-editor/default?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) 

## Support and feedback 

For any other queries, reach our [Syncfusion® support team](https://support.syncfusion.com/?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

Request new feature through [Syncfusion® feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

## License

This is a commercial product and requires a paid license for possession or use Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/34.1.29/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). You can purchase a licnense [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or start a free 30\-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 