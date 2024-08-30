**CRDP-Excel**

Note much code here other than a Excel spreadsheet with a VBA script to perform RESTAPI communications with a CRDP endpoint.

The spreadsheet contains three things:  1) VBA code, 2) sheet that performs PROTECT and REVEAL functions, 3) a sheet that performs ONLY reveal functions.
Within the first sheet, you will find the following columns:
A) CRDP URL that points to the CRDP PROTECT endpoint
B) The CRDP Protection Policy Name as defined in CipherTrust
C) A simple index column that is leverages to make each plaintext entry unique in the spreadsheet
D) Each row is sample plaintext
E) The equivalent CipherTrust as returned by CRDP
F) The CRDP Protect Profile Version (needed to return CipherText to Plaintext)
G) Blank Column Separator
H) CRDP URL that points to the CRDP REVEAL endpoint
I) A copy of the CRDP Protection Policy Name (copied from (B))
J) The name of the USER who is requesting the information (you policy will determine how REVEAL presents the result to that usser)
K) The resulting plain text

On the second sheet, you will find rows similar to the first sheet, but they are NOT linked.  This is to show that that you are truely taking ciphertext and converting it back to plain text within not spreadsheet-linking to the original ciphertext.

For your environment, you will need to change the CRDL endpoint URLS (col A and H), the name of the Protection Policy (B and I), and any user information based on your user groups (J)

Enjoy!
