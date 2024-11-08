# ZLAS_To_LAS

These two Python scripts used for ZLAS to LAS conversion and data extraction. The two Python scripts should placed inside the same folder. These scripts are only for Windows machines.

## Application

To run it, first download the file as `.zip` file and extract it, then use a command similar to the following:

```
& "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" "main.py" "path_to_ZLAS_folder"
```

* The command in a high level:

  * Run python.exe that executes main.py, and it will take in a **folder** containing ZLAS files. 

* The Python script will convert only one ZLAS to LAS at a time, extract the useful information and put to an excel spreadsheet, then delete the LAS file as it is no longer needed and would only take up extra space. Then repeat this process until information is extracted from all ZLAS files.
