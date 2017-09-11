#!/bin/bash
rm /tmp/movimientos.xls
python kutxabank.es.py -f kutxa.xls -d 10 
python xls2sql.py -f kutxa.xls
