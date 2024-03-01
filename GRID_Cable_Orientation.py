"""
Name: GRID - Cable Orientation
Autor: Andrea Sulova
Date: 27th Feb 2024
The workflow aimed at analysing cable orientation around WTG points.
The Cable Orientation is calculate from North azimuth (NAz) - clockwise direction from north.
"""

import os
import arcpy
import zipfile
import pandas as pd
from arcpy import metadata as md

#------------ Inputs
# Define input parameters fetched from the user or other sources
points_layer = arcpy.GetParameterAsText(0)
wtg_name = arcpy.GetParameterAsText(1)
x = arcpy.GetParameterAsText(2)
y = arcpy.GetParameterAsText(3)
cable_layer = arcpy.GetParameterAsText(4)
buffer_size = arcpy.GetParameterAsText(5)
output_folder = arcpy.GetParameterAsText(6)
temp_geodatabase_name = "temp.gdb"

arcpy.AddMessage(output_folder)

# Create the in-memory geodatabase
arcpy.CreateFileGDB_management(output_folder, temp_geodatabase_name)

# Set workspace to the in-memory geodatabase
arcpy.env.workspace = output_folder + "\\" + temp_geodatabase_name

# -----Check Inputs
if not os.path.exists(output_folder):
     arcpy.AddMessage("Output Folder Does not Exit, please create a folder.")
if arcpy.Exists(points_layer):
    arcpy.AddMessage("WTG Feature layer exists.")
else:
    arcpy.AddMessage("WTGFeature layer does not exist.")
if arcpy.Exists(cable_layer):
    arcpy.AddMessage("Cable Feature layer exists.")
else:
    arcpy.AddMessage("Cable Feature layer does not exist.")
    
#----- 1) Create a buffer around the WTG points
buffer_size_meters =  buffer_size + " Meters"
buffer_output= "Buffer_"+ buffer_size.replace(" ", "_")
arcpy.AddMessage("1) Establishing buffer zones around the WTG points")
arcpy.Buffer_analysis(points_layer, buffer_output, buffer_size_meters)

#----- 2) Set up field mapping for spatial join between Cables and WTG
arcpy.AddMessage("2) Spatial join between Cables and WTG buffer zones")
# Add a new text field with 255 characters
new_field_name = "FromTo"
# Check if the field already exists and delete it if found
existing_fields = [field.name for field in arcpy.ListFields(buffer_output)]
if new_field_name in existing_fields:
     arcpy.DeleteField_management(buffer_output, new_field_name)
    
# Calculate the new field using values from an existing field
arcpy.CalculateField_management(buffer_output, new_field_name, "!" + wtg_name + "!", "PYTHON")
# Set up field mapping for spatial join
field_mappings = arcpy.FieldMappings()
field_mappings.addTable(buffer_output)
# Set the merge rule for the 'Angle_between' field
arcpy.AddMessage("3) Identifying WTG names beetween the angle will be measured")
for field in field_mappings.fields:
    if field.name == new_field_name:
        field_index = field_mappings.findFieldMapIndex(field.name)
        if field_index != -1:
            field_map = field_mappings.getFieldMap(field_index)
            field_map.mergeRule = 'Join'
            field_map.joinDelimiter = ' - '
            field_mappings.replaceFieldMap(field_index, field_map)
            
# Perform spatial join
arcpy.analysis.SpatialJoin(
    cable_layer,
    buffer_output,
    "Spatial_Join",
    join_operation="JOIN_ONE_TO_ONE",
    join_type="KEEP_ALL",
    field_mapping=field_mappings,
    match_option="WITHIN_A_DISTANCE",
    search_radius="20 Meters"
)
# Perform intersect analysis
arcpy.AddMessage("4) Perform the intersect analysis between buffer zones and cable lines")
intersections_output = "Intersect"
arcpy.analysis.Intersect([buffer_output, "Spatial_Join"], intersections_output, "ALL", None, output_type="INPUT")
arcpy.AddMessage("5) Iterate through each row in the feature class to identify the names of the Start (From) and End (To) WTGs.")
new_name_start = "Start"
if not arcpy.ListFields(intersections_output, new_name_start):
    arcpy.AlterField_management(intersections_output, wtg_name, new_name_start ,new_name_start)
    
new_name_end = "End"    
if not arcpy.ListFields(intersections_output, new_name_end):
    arcpy.AlterField_management(intersections_output, "FromTo_1", new_name_end, new_name_end )

    
# Iterate through each row in the feature class using an update cursor
with arcpy.da.UpdateCursor(intersections_output, ["Start",  "End"]) as cursor:
    for row in cursor:
        split_values = row[1].split(' - ')
        if len(split_values) >= 2:
            if split_values[0] == row[0]:
                row[1] = split_values[1].strip()
                cursor.updateRow(row)
            if split_values[1] == row[0]:
                row[1] = split_values[0].strip()
                cursor.updateRow(row)
                
# List of fields to keep
fields_to_keep = [new_name_start, new_name_end, x, y]
# Create field mappings object for FeatureClassToFeatureClass_conversion
field_mappings = arcpy.FieldMappings()
# Iterate through all fields in the input feature class
for field in arcpy.ListFields(intersections_output):
    if field.name in fields_to_keep:
        field_map = arcpy.FieldMap()
        field_map.addInputField(intersections_output, field.name)
        output_field = field_map.outputField
        output_field.name = field.name
        field_map.outputField = output_field
        field_mappings.addFieldMap(field_map)
# Copy selected fields to a new feature class
output_feature_class = "Angle"
arcpy.FeatureClassToFeatureClass_conversion(intersections_output, arcpy.env.workspace,
                                           arcpy.ValidateTableName(output_feature_class, arcpy.env.workspace),
                                           field_mapping=field_mappings)
# Add new float fields to the feature class
arcpy.AddField_management(output_feature_class, "X1", "Text")
arcpy.AddField_management(output_feature_class, "Y1", "Text")
arcpy.AddField_management(output_feature_class, "X2", "Text")
arcpy.AddField_management(output_feature_class, "Y2", "Text")
arcpy.AddField_management(output_feature_class, "AngleFromNorth", "Double")
# Create an update cursor for the feature class
with arcpy.da.UpdateCursor(output_feature_class, ["SHAPE@", "X1", "Y1", "X2", "Y2", "AngleFromNorth", x, y]) as cursor:
    for row in cursor:
        line_geometry = row[0]
        start_point = line_geometry.firstPoint
        end_point = line_geometry.lastPoint
        if abs(start_point.X - row[6]) <= 5.0 and abs(start_point.Y - row[7]) <= 5.0:
            row[1] = round(start_point.X, 0)
            row[2] = round(start_point.Y, 0)
            row[3] = round(end_point.X, 0)
            row[4] = round(end_point.Y, 0)
            angle_from_north = math.degrees(math.atan2(end_point.X - start_point.X, end_point.Y - start_point.Y))
            angle_from_north = (angle_from_north + 360) % 360
            row[5] = angle_from_north
            cursor.updateRow(row)
        else:
            row[1] = round(end_point.X, 0)
            row[2] = round(end_point.Y, 0)
            row[3] = round(start_point.X, 0)
            row[4] = round(start_point.Y, 0)
            angle_from_north = math.degrees(math.atan2(start_point.X - end_point.X, start_point.Y - end_point.Y))
            angle_from_north = (angle_from_north + 360) % 360
            row[5] = angle_from_north
            cursor.updateRow(row)
# Set the workspace for shapefile and Excel file
output_shapefile = os.path.join(output_folder, "Cable_Angle_"+ buffer_size.replace(" ", "_")+".shp")
output_excel = os.path.join(output_folder, "Cable_Angle_"+ buffer_size.replace(" ", "_")+".xlsx")

# Export the shapefile and Excel
arcpy.CopyFeatures_management(output_feature_class, output_shapefile)
arcpy.TableToExcel_conversion(output_feature_class, output_excel)
arcpy.AddMessage("WELL DONE - you can check the output folder:")
arcpy.AddMessage(output_folder)
