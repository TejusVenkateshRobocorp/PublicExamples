*** Settings ***
Documentation       Template robot main suite.
Library                RPA.Excel.Files
Library                Collections



*** Tasks ***
Minimal task
    Log    Done.
    Open Workbook    MyExcel.xlsx
    
    #Below For loop Sets the formula to a single column looping through the rows 5 to 10
    
    FOR    ${row}    IN RANGE    5    11               #Looping through A5 to A10
        # Log To Console   ${row}
        Set Cell Formula    A${row}    =B${row}+10        # Example : Set A5 Formula to B5+10
    END

    ${columnsList}=    Create List    A     B     C     D      E      # Creating a list of columns to auto fill the formula
    #Below loop 
    FOR    ${idx}    ${column}    IN ENUMERATE    @{columnsList}
        IF    ${idx} == ${0}
            # No operation required in this scenario. 
            No Operation
        ELSE
            ${index}    Evaluate    ${idx} - 1 
            ${previousColumn}=   Get From List    ${columnsList}     ${index}
            Log To Console      ${column}1=${previousColumn}1+10
            Set Cell Formula    ${column}1    =${previousColumn}1+10
        END
    END
    Save Workbook 



