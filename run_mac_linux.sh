#!/usr/bin/env bash
set -e
mkdir -p output
python3 aml_reconcile.py --input ./input --output ./output/AML_Training_Full_Reconciliation.xlsx --config ./config.json
