# INP-225-Dynalite-hotel-VBA-tool
A VBA enabled excel file for Area planning

Conditions:
* Area 0~999 are reserved 
* Maximum area number can be used = 65535
* Minimum Floor offset difference = 300 (round up from 256)
* Preferred Floor offset difference should be x000 > x0000 > x00
* Every floor area number starts from x20, x0~x19 are reserved (temporary)
* Every Room occupies 20 Areas (temporary) if later this Area_per_Room changes, Reserved every floor area number (x0~x19) may changes accordingly. 
* 
