#!/bin/bash

if [ -z "$1" ]; then
    echo "Error: Please provide a source file"
    exit 1
fi

if [[ ! "$1" =~ \.xlsx$ ]]; then
    echo "Invalid file format: Currently onlt .xlsx is supported."
    exit 1
fi

source venv/bin/activate

pip3 install openpyxl

python3 main.py "$1"

deactivate
