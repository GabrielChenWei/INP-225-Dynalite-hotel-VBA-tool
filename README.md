# INP-225-Dynalite-hotel-VBA-tool
A VBA enabled excel file for Area planning

### Conditions:
* Area 0~999 are reserved 
* Maximum area number can be used = 65535
* Minimum Floor offset difference = 300 (round up from 256)
* Preferred Floor offset difference should be x000 > x0000 > x00 (x00 minimum value is 300)
* Every floor area number starts from x20, x0~x19 are reserved (temporary)
* Every Room occupies 20 Areas (temporary) if later Area_per_Room changes, Reserved every floor area number (x0~x19) may changes accordingly. 


### Variables (Public, Operators configurable, SheetConfiguration)
* Number_of_Floor
* Floor_offset_Difference
* Room_Starting_Number
  * **:o: Considering to change it to SheetRef**
* Average_Number_of_Room_per_Floor 
  * Used to generate the building prototype in SheetLayout
* Max_Room_Number_per_Floor 
  * Used to populate the numberingReferenceRow in SheetLayout.Row(2))
* Area_per_Room 
  * Fixed value (20) due to system design for now (15-Oct-2016)



### Variables (Public, System non-configurable, SheetRef)
* DipSwith_Starting_number
  * Fix to 0
* DipSwitch_per_Room 
  * Fix to 1 
* Max_Area_Number_per_Building
  * **:o: Not in use? Need to check.**
* Required_Biggest_Area_Number
  * It is calculated according to the input variable values
  * `= Number_of_Floor * Floor_offset_Difference + Last_Floor_Room_Quantity * Area_per_Room` 
* Do_Not_Show_Welcome_Info
  * True to disable Welcome message


### Roadmap
* V01~V10: Basic features: 
  * Raw file reformatting (uses formula to present the cell values) 
  * Provides formatting UI
  * Provides variable input UI
  * Provides notification during variable value input (If...Else statement on Worksheet cell directly) 
  * Define Named ranges
  * Public variable loading 
  * Cells.value calculation
  * Setup floor prototype (header: Floor numbering, Floor_offset_Difference, 1st room cells)
  * Setup 1st floor rooms (duplicate 1st room of 1st floor to Average_Number_of_Room_per_Floor)
  * Duplicates all other floors (based on 1st floor)
  * Provides **Area_Number_Overflow** checking at initial run 
    * Exit sub if Area_Number_Overflow and notification if Area_Number used up
  * Provides **Area_Overlap_Status** checking at initial run 
    * Exit sub if Area_Overlapped and notification if Area_Number used up
  * Provides Button to trigger the building prototype generation

* V11: 
  * Provides the option to disable Welcome message
* V12: 
  * Enable Floor-Room matrix adjustment
* V13: 
  * Areas Auto-assignment 


### Pending features
  * Find tuning: enables Floor-Room Matrix adjustment 
    * Using 1 Dimension Array to list the floor (floor as Index) and room quantity per floor
    
  * Auto-assignment: to 
  * SaveAs_XLSX: the final output file should be a XLSX file only to keep it light
  * Omit some room number and floor number (e.g. 4, 13 etc.)
  
