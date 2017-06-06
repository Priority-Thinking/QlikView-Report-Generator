# Report Generator

The Report Generator bridges QlikView Personal Edition (Desktop) 12.0, which can be downloaded here: http://www.qlik.com/products, and Microsoft Word to facilitate the automatic copy/pasting of charts and objects into tagged locations in the Word document while making specified selections in the QlikView engine. The intent is that a user creates a Word template and a QikView chart repository in order to run small batches of custom reports on a Windows machine.

## Getting Started

To simply run the application, open the ReportGeneratorSetup/Debug folder in the main project and download and run the ReportGeneratorSetup.msi (Windows Installer Package) file. This will create a program called "Report Generator (Active)". Download the entire project if you wish to contribute features and ideas. Don't hesistate to contact us directly!

### Understanding the Tags

Tags are placed in the Word template so the Report Generator knows where to paste QlikView objects.  

#### Chart Tags
Follow the format  
```
<CH01>
```    
The above tag will insert the chart from the QlikView document with the Object ID of CH01 at this location. You can find the Chart ID by right clicking an object, clicking Properties and finding the Chart ID field. If the Properties option is not available, you can see a full list of all charts and their IDs on a sheet by going to the Settings ribbon in the main toolbar, selecting Sheet Properties, and going to the Objects tab.

#### Looping Tags
Follow the format  
```
[FieldName]...[/FieldName]
```
These tags copy and paste whatever is between them for each possible value of FieldName. Hidden formatting symbols such as page breaks and new paragraphs will also be copied and pasted.
```
[State]<CH01>[/State]
```
The above tag will paste the chart with Object ID of CH01 for each possible value of State.

#### Selection Tags
Follow the format  
```
{'FieldName1','FieldValue1','FieldName2','FieldValue2',...}
```
These tags force specific selections into the tags.
```
<CH01{'FieldName','FieldValue'}>
```
The above tag will insert the chart with Object ID of CH01 after selecting FieldValue in the FieldName field.
```
[State{'Country','United States'}]<CH01>[/State]
```
The above tag will insert the chart with Object ID of CH01 for each possible value of State after making the selection United States in the Country field.

You can also use the selection tag by itself in the "Static Selections" text box in the application interface in order to make a universal selection for all chart objects in the Word template.

#### Quick Reference Tags
Follow the format 
```
<!QuickRefText>
```
Refer to the section below to learn how to use these tags.

#### Image Attribute Tags
Attribute tags can be added to certain chart tags in order to customize a chart image as it is inserted into Word. Chart tags have attribute tags added in the following manner:  
```
<CH01?height=2.00&width=3.55>
<CH01{'FieldName','FieldValue'}?height=2.00&width=3.55>
```
The above example sets the height to 2 inches and the width to 3.55. Specifying a single dimension will maintain the image's aspect ratio. As of version 1.1.0, only the height and width attributes can be adjusted.

#### Pivot Tables
In pivot tables any column with the label as a single asterisk will be deleted.

# Using the Application
Reference interface.png to see what the application looks like when run.

#### Set Document Paths
Browse for the Word Template (.doc and .docx) and QlikView Document (.qvf) using their respective "Browse" buttons.

#### Quick Reference Variables
If you happen to have a text object in QlikView that you reference very often add the QuickRefText (without the <>'s or !) and the Chart ID to the list. This will speed up the report generating and only works with Text Box objects in QlikView.

#### Static Selections
Use a static selection tag as mentioned above to apply a universal selection to all QlikView objects as they are copied into the Word Template.

#### Log
As the program runs, this list box is populated with the output telling the user what objects were copied successfully and where there may be errors in your Tag structure.

### Known Issues

If an error causes the program to break, your Word document will hang in limbo and prevent you from running any other actions against it. Open the Task Manager and end this process manually to continue.

### Prerequisities

To contribute, you must install the QlikView and Microsoft Word references to your Visual Studio project. Refer to the Built With section for information on how this program was built and tested.

## Example

RG-Example.doc  
We find it helpful to turn on the hidden characters in Word (Crtl + Shift + 8)

Prescription Tracker.qvw  
This document is included when you download QlikView 12.0. Go to the Program Files, Examples folder, then Data folder to find this document.

Target these two files with the Report Generator and click Generate to watch it run. 

## Deployment

If you wish to simply use the program, download and run the setup utility in the Setup.zip file.

## Built With

Microsoft Visual Studios 2015 Community - Windows Forms  
Built and Tested on Windows 8.1 OS  
Microsoft Word 15.0 Object Library Version 8.6  
QlikView 12.0 Type Library Version 12.0

## Compatibility and Testing

This project was built and tested with the following software versions:

QlikView 11.20, 12.0  
MS Word 2013, 2016  
Windows 8.1 Pro, 10 Pro  

## Contributing

Feel free to contact tkendrick@prioritythinking.com if you are interested in contributing to this project.

## Versioning

We use semantic versioning. Reference tags and description for version number.

## Authors

<strong>Priority Thinking Team</strong>  
Tim Kendrick  
John Murray  
Mitali Ajgaonkar  
Grant Parker  
Qasim Ali  

## License

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

## Acknowledgements

Rochester Institute of Technology  
https://community.qlik.com/  
www.prioritythinking.com
