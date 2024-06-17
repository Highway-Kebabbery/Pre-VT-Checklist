# <div align="center">Pre-VT Checklist</div>


## Description
### Overview


This project was scrubbed of any proprietary information and I built new databases with dummy data using SQLite for the purpose of demonstration.

## Features

I created this Excel-VBA application for use in the process of validating data migrations. The data migration flow was as follows: old production environment > new development environment > new QA/stage environment > new production environment. Data were spot-checked in the new development environment, but all data had to be inspected in the QA/stage environment during the formal review. Prior to testing objects in the QA/stage environment, it was necessary to determine that all objects and supporting objects in each batch met several criteria across three databases (old production environment, new development environment, and new QA/stage environment).

Given that a batch may contain hundreds of supporting objects across dozens of tables spread over three databases, this tool not only massively reduced the time it took to make sure objects were ready to begin testing, but it also massively reduced human error and oversight. The alternative to this tool was to perform over 30 queries manually, to evaluate data manually in Excel, and to hope that nothing was missed along the way. Even the largest batches were analyzed within 30 seconds using this tool, which represents a steep decrease from my estimate of two hours per batch split between performing the manual review and the rework that was avoided by using this tool.

Before testing could begin, it was necessary that all objects and all associated supporting objects in each batch:

* Exist in the QA/stage environment
* Be newer in the new development environment than in the old production environment
* Be newer in the new QA/stage environment than in both the new development and old production environments
* Be marked as REMOVED='F' in all three databases
* Be marked as ACTIVE='T' in all three databases

In order to achieve this, the tool performs several key actions:
* A single list of parent objects, representing one batch, is pasted into column A by the end-user. The tool connects to three databases and then uses that list of parent objects to locate all related supporting objects in LabWare LIMS and check every object for all of the criteria listed above in every database.
* For analysis objects, the tool also uses the batch_link of the analyses on the current batch to locate any other analyses (parent-level objects) that are related through a batch_link and queries them, as well, to ensure that no analysis on the current batch is adjacent to any other analysis that may not be ready for testing.
* For analysis objects, the tool pulls all calculations related to the analyses on the batch and scrapes them for instances of subroutines defined in LabWare's SUBROUTINES table so that it can validate those subroutines alongside the other supporting objects.
* Finally, any issues are consolidated into one status cell per object that presents users with any obstacles to testing the object in question.

## How to Use
### Software Requirements
* Microsoft Windows operating system (designed and implemented on Windows 10 Enterprise and Windows 11 Pro; other versions may work)
  * Not built for MacOS or Linux systems
* Microsoft Excel
* Microsoft ActiveX Data Objects 6.1 Library (other versions may work)
* SQLite ODBC Driver available <a href="http://www.ch-werner.de/sqliteodbc/" target="_blank">here</a>. (Choose Win32 or Win64 based on the bitness of your Microsoft Office installation.)

### Instructions
1. Install the required SQLite ODBC Driver linked in [Software Requirements](#software-requirements).
2. Download the latest release of the project [here](https://github.com/Highway-Kebabbery/Pre-VT-Checklist/releases).
3. Office may block the content of the application because it contains macros. Right-click "Pre-VT Checklist.xlsm" > select "Properties" > select "Unblock" at the bottom of the "General" tab if the option appears.
4. You may need to select the option to "Enable Macros" in a yellow ribbon under the main menu of the Excel file.
5. Read the instructions at the top of the spreadsheet.
6. Copy an assortment of object names from __EITHER__ the ANALYSIS or PRODUCT test objects available in cells N28:N45. Paste the objects into the golden cells of column A. For now, you __MUST__ paste the first object into cell A7. This will be fixed later.
7. Select the appropriate object type in cell A4.
8. Click "Browse."
9. Check status messages in column L.

## Technology
* **Visual Basic for Applications (VBA):** I learned VBA for this project because it's the internal language for Microsoft Office applications.
* **Microsoft Excel:** I used Microsoft Excel because it's useful for displaying and manipulating data.
* **Oracle SQL (Previously):** The project was originally designed for use with an Oracle database.
* **SQLite:** The demonstration version of this project required dummy data to prove its functionality. SQLite was the obvious choice due to its ease of use.
* **SQLite ODBC Driver:** Available <a href="http://www.ch-werner.de/sqliteodbc/" target="_blank">here</a>. This was required to establish a connection between the application, which uses ADODB objects, and the SQLite database.

## Collaborators
Thank you to Daniel Guichard for authoring the class module that establishes the database connection. This module required only minor modifications for my purposes.


## What I Learned
This project presented many challenges and it represents a steep increase in complexity over my previous projects. I knew the fundamentals going into this project, but it gave me my first real opportunity to use everything that I knew as well as so much more. I'll surely forget all of the minutiae by the time I interview for a job, so here's what I still remember having learned as of this writing:

**VBA**:
* First and foremost, I learned an entirely new language just to complete this project: Visual Basic for Applications (VBA). Much of what I learned entails little more than implementing fundamental concepts in a new language.
* By the last module (ColTitles, which eliminated magic numbers) I learned how to set up a properly-encapsulated class in VBA. I still find the implementation of classes in VBA to be a little odd (the inability to access private class attributes through Me.attributeName??).
* I learned how to stream a CLOB to a string in VBA. Although this is no longer required in the demonstration version of this application given that SQLite doesn't support CLOBs, the code still functions and was left in place to demonstrate what I did in the real version of the application.
* I learned how to create ADODB recordsets and interact with them using VBA in Microsoft Excel.
* I got lots of practice designing selection statements that short-circuit in my code.
* I learned how to declare variably-sized arrays in VBA.
* I learned how to (re-)format a spreadsheet using VBA (which is amazing because nobody seems to know how to paste text-only in my beautiful works of art).

**SQL/VBA:**
* I gained a lot of practice writing flexible SQL query templates that take variable input.
* I learned how to connect a VBA application's back end to both an Oracle SQL server (in the original implementation) and to a SQLite file using an ODBC for SQLite driver (in the demonstration version of the application)
* I learned that SQLite does not support the CLOB data type, and furthermore that you have to check the database structure to see the actual data type, because "SELECT type_of() FROM <table_name>" seems only to return the data type that SQLite thinks is appropriate given its flexible data typing.

**General:**
* I reinforced my ability to break down a problem, design a programmatic solution to that problem, implement a minimum viable product, and finally to implement additional features in the application.
* I got very familiar with the database structure of LabWare LIMS V8.
* I learned how to run macros from an ActiveX button in Excel.


This was an extra-curricular project that was secondary to, but intended to support, my primary task of testing the data we migrated. As such, there are things I would do differently under less constrained circumstances.  The main thing I would change about this project is to go back and rewrite my classes to be truly encapsulated and accessible only through public getter, setter, and deleter functions. The class I wrote last, ColTitles, meets these criteria, but I was under a very tight deadline and still learning VBA when the rest of the project was written so I wrote what worked and moved on. I understand how these class modules can be improved. I would also update the functionality checking to see if all parent objects were located in the D3 database. This represents a way to make the application more robust against "incorrect" use.

### Why I'm Showcasing this Work
This project not only shows that I understand the fundamentals of writing software, but I firmly believe that it also shows that I am solidly above the level of a junior software engineer and, as such, could be hired as a mid-level engineer.

This work displays:
* My grasp of fundamental programming concepts, including object oriented programming.
* My knowledge of SQL.
* My willingness and ability to learn an entirely new language for a project (namely: VBA).
* My ability to add sensible comments.
  * I'll note that some additional commentary was added because I always knew this would be a portfolio piece and felt they were appropriate in that context to demonstrate *why* I made decisions I made. This will help me more thoroughly discuss my work in the future when reviewing it for interviews.
* My ability to program without magic numbers.
* My ability to design programs that short-circuit to optimize performance.
* My ability to build the back end of an application that communicates with a database.
* My ability to construct the front end of an application. Though it is admittedly very simple given that it's a spreadsheet, it's designed to be as easy to use and understand as possible.
* My understanding of relational databases and the ability to construct one when necessary, albeit very simple.
* As evidenced by this README.md: my ability to communicate clearly.

