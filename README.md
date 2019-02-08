## This is a macro that performs a series of find/replace operations on a collection of MS word documuments.

#### Setup

The setup just involves organizing the files as shown in SetupDiagram.png:
* FindReplaceMacro.dotm in working directory
* Have lookuptable.csv alongside the macro
* Make a directory titled **Files** alongside the macro and look-up table
* **Files** directly contains all .docx files on which you wish to execute the find/replace operations
* The look-up table file and file directory should not be renamed from this schema unless you want to edit the macro yourself

![Alt](/SetupDiagram.png "Setup Diagram")

#### Run

Openning the macro will appear to simply open a word document.  To run:
* You will need macros enabled (a security setting)
* You will need the developer tab enabled
* Click on **Macros** button on **Developer** tab
* Select *findReplaceAllFiles* and press **Run**
* **FilesWithSubs** will now contain the files with substitutions made. If the file name was in the substitution table, the file will be renamed according to the substitution.