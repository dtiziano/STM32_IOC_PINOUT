# STM32_IOC_PINOUT
This code can be used to read a .ioc file created in CubeMX and to read a kicad schematic file and export the pinout to an excel sheet. 

In the excel sheet one can view differencies between kicad and ioc marked in red. 

This can be helpful with MCU's which habe many pins. 

If the pin naming functionality within CubeMX is used, the names will also appear in the excel sheet. 

# Usage:

## 1. Generate the .ioc 
Using Cube IDE or Cube MX, generate the .ioc file for your project

## 2. Generate a kicad schmeatic.
Using kicad, generate a kicad schematic that contains the mcu.

## 2. Generate the .xlsx 
Using the provided Jupyter Notebook ```jp_main.ipynb``` and substituting the ```file_path``` with the path to your ```.ioc``` file, ```schematic_file``` with the path to your ```.kicad_sch``` file, ```lib_file``` with the path to your ```.kicad_sym``` file (from the library) and ```unit``` the number of the unit of the symbol to be used.

The script then generates an excel file.

# For issues or functionality improvement use the Pull Request or Issue section on [github.com](https://github.com/dtiziano/STM32_IOC_PINOUT)

