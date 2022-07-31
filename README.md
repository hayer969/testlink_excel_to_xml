# Testlink excel to xml  
Script for inject text from excel to exported testcase xml file  

## Script usage:
- export ONE testcase from testlink as xml file
- create excel file (xls, xlsx) where first column is "Actions"  
  second column is "Expected Results". File should NOT contains  
  any headers or other information, data should be on first sheet  
- run this file with first parameter path to excel and second path  
  to xml. Actual command see down below.  
- Export xml file back to the testlink. You have options for already  
  existed cases: update latest version (prefered), create new version  
  and so on.  

## Requirements:
```pip install pandas openpyxl```

## **ATTENTION**: Next command will change xml file  
if you want to preserve original file, please copy it manually.  
```python testlink_xml_inject.py {path to excel file} {path to xml file}```
