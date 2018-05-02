# Open-Rolodex
A little software for quick acces to your contact

The current design responds to a specific need. Do not expect something flexible. :)

Potentially more complicated than necessary, but it comes from a first exploration of python

Make for windows.

main.py : Start the application

setup.py: used by cx_Freeze to create a standalone executable file for windows 

### List of Dependencies:
(install with 'pip install mypackage')
*  lxml
*  Pillow

#### optionnal:
*  cx_Freeze (to create a standalone executable(exe) file)

### In case of a standalone executable
You need to install *Redistribuable Visual C++ pour Visual Studio 2015*
You should find it [here](https://www.microsoft.com/fr-fr/download/details.aspx?id=48145) (x86 or x64, depend of your version of python and the setup. Test was with a x86)
