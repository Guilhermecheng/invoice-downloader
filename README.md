# invoice-downloader
### Python script to download warranty invoices from Dealernet

Using data obtained from an excel file with the invoice numbers needed to download, the script goes service order to service order downloading each of the invoices.

### excel.py
Script responsible to get the data searched, and pass it to the main script.
Data comes in the following format:

| OS_Numero | Empresa_Nome | NotaFiscal_Numero  | OS1 | OS2 |
| :---:   | :-: | :-: | :-: | :-: |
| 2303373 | BRG FLORIANOPOLIS 0003 | 476   /723    | 476 | 723 | 


Where:
OS_Numero - service order number
Empresa_Nome - Dealer name
NotaFiscal_Numero - invoice numbers from the report
OS1 - invoice number 1
OS2 - invoice number 2

OS1 and OS2 are columns created to ease our search in the script.

### dealernet.py
The main script.
It goes on the following order:
1. login in system
2. search for the service order in the system, and its dealer
3. in the SO, look for each invoice number, OS1 then OS2, enter it, and download the invoice pdf file. 

