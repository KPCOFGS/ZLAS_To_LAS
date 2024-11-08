import arcpy
import sys

def convert_zlas_to_las(in_zlas, target_folder):
    file_version = "1.4"
    point_format = 6
    compression = "NO_COMPRESSION"

    # Perform the conversion
    arcpy.conversion.ConvertLas(in_las=in_zlas, target_folder=target_folder, 
                                 file_version=file_version, 
                                 point_format=point_format, 
                                 compression=compression)
    print(f"Successfully converted {in_zlas} to LAS.")

if __name__ == "__main__":
    # Get command-line arguments
    in_zlas = sys.argv[1]
    target_folder = sys.argv[2]

    # Run the conversion
    convert_zlas_to_las(in_zlas, target_folder)
