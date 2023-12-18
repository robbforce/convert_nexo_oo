Based on convert_nexo.vbs v1.3 by droblesa (03/18/21), https://community.accointing.com/t/nexo-integration/95/62

These vbscripts were written for LibreOffice and can be used to convert nexo.io transaction history to various templates.

INSTRUCTIONS:

* Download and install LibreOffice: https://www.libreoffice.org/
or install via Chocolatey: https://community.chocolatey.org/packages/libreoffice-fresh
* Downlod your transaction history from nexo.io
* Drag and drop your nexo csv file onto the vbscript to convert.

NOTES:
* As of 2023-12-11, Coinpanda's parser appears to choke on CSV files with more than 100 rows, so I wrote a batch loop that will save the output to smaller CSVs. It's set to 100 rows, but simply set iBatchCount at the top of the script to change it. The script also outputs the combined converted CSV.
