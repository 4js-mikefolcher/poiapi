# FGL wrapper for the Apache POI Java Library

## Description

This example program wraps the Java Apache POI libraries in 4gl libraries
so that they can be called by a 4gl programmer.

Currently, only the ability to create a spreadsheet is supported, but the plan is add the ability to import and export to
different types of MS Office files.

The Apache POI libraries are used to interact with Excel documents.

## Usage

You will need to download the Java Apache POI libraries from:

https://poi.apache.org/

In the fgl_apache_poi node, set the POI_HOME variable to where you have 
downloaded the Apache POI Libraries

As the versions increase you may need to alter the CLASSPATH variable set
in the same node.

Currently I use the value for CLASSPATH in my Genero Studio environment ...
``
$(CLASSPATH);$(POI_DIR)/poiapi-4js-5.2.3.jar;$(POI_DIR)/log4j-core-2.19.0.jar
``
... the jar files correspond to the files in the javadeps directory of the repo.


## Test Program

### fgl_excel_api_test

Creates three different Excel spreadsheets using both the TSpreadsheet and TSpreadsheetXtend API's.

## Notes

The file fgl_excel.4gl originated from Reuben's GitHub repo fgl_apache_poi (https://github.com/FourjsGenero/fgl_apache_poi).  I decided to focus on exporting to excel and I wanted to make the interface cleaner, so I built this library on top of his fgl_excel.4gl.  I needed to make some changes to support the newer version of POI and to implement some options that were not supported in the original version.
