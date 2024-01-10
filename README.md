# STM32_IOC_PINOUT
 This code can be used to read a .ioc file created in CubeMX and export the pinout to an excel sheet. 

 This can be helpful with MCU's which habe many pins. 

 If the pin naming functionality within CubeMX is used, the names will also appear in the excel sheet. 

 # Usage:

 ## 1. Generate the .ioc 
 Using Cube IDE or Cube MX, generate the .ioc file for your project

 ## 2. Generate the .xlsx 
 Using the provided Jupyter Notebook ```jp_main.ipynb``` and substituting the ```file_path``` with the path to your ```.ioc``` file, the excel sheet corresponding to your file can be created.

 # For issues or functionality improvement use the Pull Request or Issue section on [github.com](https://github.com/dtiziano/STM32_IOC_PINOUT)

