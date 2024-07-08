# Requirement
- An Excel Table containing columns PM Name, Project Status.
- Hundreds of Projects distributed between the PMs
- Need to measure the Completion % of each PM
- Project Status can be in one of Scheduling, Development, UAT, Complete, Hold, Blocked etc...
- Weights assigned to each, lowest 1 (Scheduling) to highest 4 (Complete)
- Target Score for each PM is 4 * Total count of projects assigned
- Current score for each PM is Sum of (Weight*Count of Projects in that Status). Example if a PM has 2 Complete, 1 in Development and 3 in UAT, score is (2*4)+ (1*2) + (3*3) = 19. Target Score is 6*4 = 24
- Completion Score % is (19*100)/24
- Completion Score is monitored weekly
- Need a script to compute the Completion score.
## Source Table
![image](https://github.com/ravichinni/excel-script-project-status-summary/assets/10705784/645006a1-c73c-48d9-9197-77c018efc1c1)

## Target Table
![image](https://github.com/ravichinni/excel-script-project-status-summary/assets/10705784/c64a484c-ae4c-40f4-9cb7-1b339e3bbfb9)

## Sample file
See the sample file https://github.com/ravichinni/excel-script-project-status-summary/blob/main/Project%20Status%20Tracker.xlsx

# Solution - Used GPT-4o to generate the code
**These are the prompts**
- Prompt #1: You are a Expert Developer on Office Scripts particular on Excel Script
- Prompt #2:
Below are the requirements for the script.
1. There's a table in the Excel workbook by name "Table1". This table lists all the projects being executed in the company. The table contains two columns "PM", "Project Status". PM represents the Project Manager managing the project. Current status of the project could be one of many values, and the current status is captured in the "Project Status" column.
2. There's another table by name "Table11", intent of this table is track progress. The Rows are the PMs, and the columns are "Project Score on the Date". The Project Score is the count of all Projects a PM handles. 
3. Generate an Excel script to:
4. Add a row for each PM in the Table11 if it doesn't exist already.
5. Add a column to the table with column name as Current Date in the format "DDMMMYY". Example: 08JUL24. The column should be added if it isn't present already.
6 The row value for the column should contain a Project Score. Project Score is defined as the count of the projects the PM has.
- Prompt #3: Got a compilation error "See line 30, column 9: Office Scripts cannot infer the data type of this variable. Please declare a type for the variable."
- Prompt #4: Another error "[30, 41] Property 'getColumnCount' does not exist on type 'Table'."
- Prompt #5: Next error "[52, 31] Property 'getRangeBetweenHeaderAndLastRow' does not exist on type 'Table'. Did you mean 'getRangeBetweenHeaderAndTotal'?"
- Prompt #6: next error "[47, 29] Type 'IterableIterator<[string, number]>' is not an array type or a string type. Use compiler option '--downlevelIteration' to allow iterating of iterators."
- Prompt #7: Please add debug statements to the code at appropriate locations
- Prompt #8: Now lets change the logic to calculate the Project Score. Each of the Project Statuses carry a weight as stated below.
Complete is 4
UAT is 3
Development, In Progress is 2
Rest all statuses is 1
NotApplicable is 0
For each PM multiply the weight by the number of projects with corresponding status and sum it up.
- Prompt #9: Also add a column called "Target Score". For each PM, this is total number of projects multiplied by the weight for Complete status.
This needs to be re-calculated every time the program runs.
- Prompt #10: Also add a column with the name format "Score%-DDMMMYY". This is similar to "Score-DDMMMYY", but just expressed as a percentage of "Target Score"

