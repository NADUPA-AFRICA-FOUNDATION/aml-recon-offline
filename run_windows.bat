@echo off
if not exist output mkdir output
python aml_reconcile.py --input .\input --output .\output\AML_Training_Full_Reconciliation.xlsx --config .\config.json
pause
