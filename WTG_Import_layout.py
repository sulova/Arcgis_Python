

"""
Name: Import WTG layout 
Autor: Andrea Sulova
Date: 4th Dec 2024

Description: This script imports data from an Excel file into a geodatabase
as a feature class, adds coordinates to the attribute table, extracts metadata
from an inventory file, establishes alias names, and exports the feature class
to Shapefile and DWG formats.
"""

import os
import arcpy
import zipfile
import pandas as pd
from arcpy import metadata as md

#------------ Inputs
# Define input parameters fetched from the user or other sources
file_path = arcpy.GetParameterAsText(0)

sheet = arcpy.GetParameterAsText(1)

ID_Column = arcpy.GetParameterAsText(2)

X_Column = arcpy.GetParameterAsText(3)

Y_Column = arcpy.GetParameterAsText(4)

spatial_reference = arcpy.GetParameterAsText(5)

output_gdb = arcpy.GetParameterAsText(6)

output_fc = arcpy.GetParameterAsText(7)

data_inventory = arcpy.GetParameterAsText(8)

arcpy.env.workspace = output_gdb

# -------- Import excel files to Dataframe

# Split the file path into directory and file name
directory, filename = os.path.split(file_path)

# Read dataframe = excel file
df = pd.read_excel(directory, sheet_name = sheet)

# Remove space in columns at the end of name
df.columns = df.columns.str.rstrip()

# Create a new feature class in the geodatabase
fc_path = os.path.join(output_gdb, output_fc)

# Delete the feature class if it exists
if arcpy.Exists(fc_path):
    arcpy.AddMessage("Feature class already exit in Geodatabase")
    # exit()
    arcpy.management.Delete(fc_path)

# Check if the feature class does not exist and create it
if not arcpy.Exists(fc_path):
    arcpy.CreateFeatureclass_management(output_gdb, output_fc, "POINT", spatial_reference= spatial_reference)

    # Define fields from Excel columns and their data types
    field_mappings = [("ID", "TEXT"), ("Point_X", "DOUBLE"), ("Point_Y", "DOUBLE")]

    # Add fields to the feature class
    for field_name, data_type in field_mappings:
        arcpy.management.AddField(fc_path, field_name, data_type)

    # Iterate through Excel data and insert rows into the feature class
    with arcpy.da.InsertCursor(fc_path, ["SHAPE@", "ID", "Point_X", "Point_Y"]) as cursor:
        for index, row in df.iterrows():
            # Get X and Y coordinates from Excel (assuming columns "Longitude" and "Latitude")
            x, y  = row.loc[X_Column], row.loc[Y_Column]
            # Create a point geometry
            point = arcpy.Point(x, y)
            # Create a new row with point geometry and other attributes
            cursor.insertRow([arcpy.PointGeometry(point, spatial_reference), str(row[ID_Column]), str(row[X_Column]), float(row[Y_Column])])

arcpy.AddMessage("The feature class in geodatabase is created successfully")


# -------- Adding  coordinates to the attribute table

# Add geometry attributes
geom_props = "POINT_X_Y_Z_M"
arcpy.AddGeometryAttributes_management(fc_path, geom_props , "", "", spatial_reference)

# Rename the fields POINT_X and POINT_Y to X_COORD and Y_COORD
arcpy.AlterField_management(fc_path, "POINT_X", "X", "X")
arcpy.AlterField_management(fc_path, "POINT_Y", "Y", "Y")

# Set the output coordinate system
out_sr = arcpy.SpatialReference(4258)   # ETRS1989 - 4258
geom_props = "POINT_X_Y_Z_M"
arcpy.AddGeometryAttributes_management(fc_path, geom_props , "", "", out_sr)

# Rename the fields POINT_X and POINT_Y to X_COORD and Y_COORD
arcpy.AlterField_management(fc_path, "POINT_X", "X_ETRS", "X [ETRS 1989]")
arcpy.AlterField_management(fc_path, "POINT_Y", "Y_ETRS", "Y [ETRS 1989]")
arcpy.AddMessage("XY coordinates are added successfully")


#--------- Metadata extracted from the data inventory file

# Read the Excel file into a DataFrame
df = pd.read_excel(data_inventory)

# Filter the DataFrame based on the 'Full Name' column containing the search string
filtered_df = df[df["Full Name"] == output_fc]

# Check if the search string was found in the DataFrame

if not filtered_df.empty:
    # Get the text from text_column in the same row
    imported_title = filtered_df.iloc[0]["Full Name"]
    imported_summary = filtered_df.iloc[0]["Summary"]
    imported_tags = filtered_df.iloc[0]["Tags"]
    imported_Description = filtered_df.iloc[0]["Description"]
    imported_Credits = filtered_df.iloc[0]["Credits"]
    imported_Date = str(filtered_df.iloc[0]["Date Created"])
    imported_Description = imported_Description + "\n" + imported_Date
else:
    arcpy.arcpy.AddMessage("Please add information into the DATA INVENTORY File (Excel)")
    arcpy.arcpy.AddMessage("Name of feature class should be the same as in the data inventory (column Name)")
    # exit()

# Create a new Metadata object and add some content to it
new_md = md.Metadata()
new_md.title = imported_title
new_md.tags = imported_tags
new_md.summary = imported_summary
new_md.description = imported_Description
new_md.credits = imported_Credits

# Assign the Metadata object's content to a target item
tgt_item_md = md.Metadata(fc_path)

if not tgt_item_md.isReadOnly:
    tgt_item_md.copy(new_md)
    tgt_item_md.save()

arcpy.AddMessage("Adding metadata is completed successfully")


#-------  Establishing Alias Names
    
# Get the describe object for the feature class
desc = arcpy.Describe(fc_path)
current_alias = desc.aliasName
arcpy.AddMessage(current_alias)

# Create a new alias name by replacing underscores with spaces
new_alias = output_fc.replace("_", " ")

try:
    # Set the new alias name for the feature class
    arcpy.AlterAliasName(fc_path, new_alias)
    arcpy.AddMessage("Alias name changed successfully.")

except Exception as e:
    arcpy.AddMessage("An error occurred during creating alias name")


#------  Export this feature class to Shapefile
    
# Create the full path to the folder where the shapefile will be save
directory = os.path.dirname(directory)
shp_folder = os.path.join(directory, output_fc)
shp_file = output_fc + ".shp"
output_shp = os.path.join(shp_folder,shp_file)
arcpy.AddMessage(output_shp)

# Check if the folder already exists before creating it
if not os.path.exists(shp_folder):
    # If the folder does not exist, create it and convert the feature class to a shapefile
    os.makedirs(shp_folder)
    arcpy.FeatureClassToFeatureClass_conversion(fc_path, shp_folder, shp_file)
    arcpy.AddMessage("Shapefile is created successfully and saved in this folder:")
else:
    # If the folder already exists, inform the user
    arcpy.AddMessage("SHP Folder already exists.")


#------ Feature class to a DWG file for CAD
    
# Set up paths and file names
dwg_folder = os.path.join(shp_folder+ '_DWG')

# Output DWG file
dwg_output = os.path.join(dwg_folder, output_fc +'.dwg')

# Check if the output path exists, if not, create it
if not os.path.exists(dwg_folder):
    os.makedirs(dwg_folder)
    arcpy.conversion.ExportCAD(output_shp, "DWG_R2018", dwg_output, False, False)
    arcpy.AddMessage("DWG file is created successfully and saved in this folder:")
else:
    arcpy.AddMessage("DWG Folder already exists.")
