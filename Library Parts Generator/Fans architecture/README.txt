%================================================================================
%================================================================================

                   ==      ==    ===      ===    ==
                     ==  ==      == ==  == ==    ==
                       ==        ==   ==   ==    ==
                     ==  ==      ==        ==    ==
                   ==      ==    ==        ==    ========

  ==             ==  ========  =========  ==========  =======  ========
   ==           ==   ==    ==     ===         ==      ==       ==    ==
    ==   ==    ==    ========     ===         ==      =======  ========
     == == == ==     ==  ==       ===         ==      ==       ==  ==
       ==   ==       ==    ==  =========      ==      =======  ==    ==

%================================================================================
%================================================================================


Instructions:

1. Save the Fan_Master.xml and Fan_Write.xlsm into your own directory

2. Open the "Fan_Write" Excel File

3. Enable macros

4. Go to Sheet 4 (Named "File_Location")

5. Enter the path name where you saved Fan_Master.xml into cell B1
   (e.g. C:\Users\JDoe\Documents)

6. Enter the path name where you want to output the xml files into cell B2
   (e.g. C:\Users\JDoe\Documents\xml Files)

7. Enter the path name where the Libraries folder in FloTHERM is located at
   (e.g. C:\Program Files (x86)\MentorMA\flosuite_v12\flotherm\flocentral\
    Libraries)

8. Select whether you want to create xmls for all models in the All_Models tab
   or selected models in the Selected_Models tab using the drop down menu
   in cell B4

9. Go to the view tab and click on Macros

10. Select the Fan_Write macro and click run

11. Select the Compile macro and click run

12. Select the Create_Libraries macro and click run

13. Open FloTHERM®

14. Click Macro in the toolbar and then click play FloSCRIPT

15. Run the Create_Libraries.xml in the folder specified in step 7

16. Run the vendorName.xml found in the vendor's folder to execute
    all the xmls within that folder

17. Double click on the part file in the library to load it into the project



%================================================================================
%================================= General Tips =================================
The values used in FloSCRIPTs are case sensitive.
For example, 'clockwise' and 'Clockwise' would be interpreted differently

The compile subroutine compiles all the files listed in the spreadsheet,
not all the files in the folder. So ensure the files you want are listed

Please avoid adding/removing columns as the code runs specifically for the
current number of columns

Please avoid moving cells from place to place as the VBA code references
these specific cells

If running a FloSCRIPT for the same part name after the first time
(this could be if the data entered was wrong and you wanted to re-create
the part), please ensure the old part file is deleted as running the
FloSCRIPT will not automatically overwrite the library part with the same name

Make sure you are using a new instance of FloTHERM before running the FloSCRIPTs



%================================================================================
%===========================Uses of each Sheet===================================

%=============================== Sheet 1 ========================================
Stores all the models for all the different vendors
Note: Cannot have gaps between each data entry otherwise Macro won't work

%=============================== Sheet 2 ========================================
A dynamic sheet which can be re-written on
Used for outputting specific xmls for models which for example, needed editing
Note: Cannot have gaps between each data entry otherwise Macro won't work

%=============================== Sheet 3 ========================================
Stores the vendors' names to create folders corresponding with those names

%=============================== Sheet 4 ========================================
Stores file path to read Master file from
Stores file path to output xmls to
Used to select running Macros for sheet1 (All_Models) or sheet2 (Selected_Models)
Used to select FloTHERM® version no.



%================================================================================
%============================Use of each Macro===================================

%=============================== Fan_Write ======================================
Outputs each model no. as an individual xml file

%================================ Compile =======================================
Compiles all the xmls for all models listed in the spreadsheet tab into a single
file for each vendor

%=========================== Create_Libraries ====================================
Creates libraries in FloTHERM®'s installation folder so it can find the file
location when saving the PDMLs

