#!/bin/bash

# Directory to convert xls to xlsx
DIRECTORY=~/Documents/logs/kaspersky_av

# Current date to be searched in files
DATE=$(date +"%-m-%-d-%Y")

# Check if directory exists
if [ ! -d "$DIRECTORY" ]; then
	echo "Directory $DIRECTORY does not exist"
	exit 1
fi

# Command to convert xls to xlsx
COMMAND="libreoffice --convert-to xlsx --outdir $DIRECTORY {} --headless"

# Check if libreoffice is installed
libreoffice --version > /dev/null
libreoffice_exit_code=$?

if [ $libreoffice_exit_code -ne 0 ]; then
	echo "libreoffice not installed"
	echo "Install libreoffice and try again"
	exit 1
fi

# Search for files matching the pattern and convert to xlsx
find "$DIRECTORY" -name "Report-EDR*$DATE*.xls" -exec $COMMAND \;
find_exit_code=$?

# Check if find command was successful
if [ $find_exit_code -ne 0 ]; then
	echo "xls to xlsx conversion failed with error code : $find_exit_code"
	exit $find_exit_code
else
	# Rename the converted xlsx files
	find "$DIRECTORY" -name "Report-EDR(*$DATE*).xlsx" -exec mv {} "$DIRECTORY/Report-EDR_$DATE.xlsx" \;
	find "$DIRECTORY" -name "Report-EDR-Optimum(*$DATE*).xlsx" -exec mv {} "$DIRECTORY/Report-EDR-Optimum_$DATE.xlsx" \;

	# Remove the old xls files
	rm "$DIRECTORY"/*.xls
fi