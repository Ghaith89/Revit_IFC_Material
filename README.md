# Revit_IFC_Material
The problem when importing IFC projects to revit is that it is not possible to change their material because each geometry of the imported family will have his own material in the material library. This will result hundrets of materials in the material library which have the same texture and the same color key.
This program is done to categorize imported IFC materials in colorkeys and adjust them using Excel tabel. It supposes to help controling the IFC imported project's materials, their appearances and names based on an excel sheet conrolled by the user.
**Coloring Excel Script**

This script is responsible for picking the coplor keys from the imported IFC object's materials and add them to a saved excel file.
As a resusult the colums 2, 3, 4 will be colored with an RGB color value added to the cells.
Column 5 will be changed manually by the user and it includes real material names corresponds to the colors.
This will be later used to change the ifc material names in the second script.
Column 1 will include the material appearances. This will correspond to the material appearances in Revit 
and the names are taken from the material's manager.
The file shall be saved in the name of the file that will be read by the other script.

**Reading material**

This script is responsible for reading the excel file generated by the Coloring Excel Script.
It supposes to read it and assign the apearances and the names to the ifc materials of all geometries in the scene.
