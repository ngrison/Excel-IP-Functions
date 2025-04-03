# Excel IP Functions

## A Suite of Excel Tools to manipulate IPv4 and IPv6 addresses in Microsoft Excel.

These tools are a useful tool for anyone who needs to interact with IP addresses in any forms. They can help you with such queries:
- What is the network address, subnet mask or broadcast address?
- What is the next or previous subnet?
- I need a list of subnets in a summary?
- Is an address or network in a multicast or private range?

These tools help you answer all of those queries and many, many more.

There are now two versions of the tools available.

## NEW IN 2025: IP Formulas

I have now developped a formula based suite of functions to manipulate IPv4 addresses.
There are two major benefits to the formula based suite:
- They run on Excel Online
- They run in standard .xls workbooks, no need for a .xlsm macro enabled workbook

They do support most features from formatting to sorting and subnetting.

The two main restrictions:
- They required Excel 365
- They only support IPv4

## Classic VBA Based Functions

If you do need to manipulate IPv6 addresses, or use an older version of Excel, my classic VBA based fucntions are still available.

They do support IPv4 and IPv6.

### Utilisation

When installed and enabled in the workbook, all functions become available in formulas and begin with "ip". Just type"=ip" in a cell and you will see the list of functions.
The functions have specific options to format the output.

Some examples of available function:
ipAddress: Get the ip address of a ip/prefix pair
ipHostX: Get the Xth host address from the beginning of a subnet
ipHostY: Get the Yth host address from the end of a subnet
ipHostFirst: Get the first host address in a subnet
ipHostPrev: Get the previous host address in a subnet
ipHostNext: Get the next host address in a subnet
ipHostLast: Get the last host address in a subnet
ipHostCount: Get the number of host addresses in a subnet
ipSubAddress: Get the network address
ipSubMask: Get the subnet mask
ipSubBroadcast: Get the Broadcast address
ipSubPrev: get the previous subnet
ipSubNext: Get the next subnet

The modules include countless more functions and format options so I highly recommend using my instruction XLSM file to view all the functions and see how they work:
[Excel IP Functions.xlsm](VBA/Excel%20IP%20Functions.xlsm)

### Installation

Manual Installation:
Create a new macro-enabled workbook, open the VBA environment and import the .bas and .cls files in the VBA folder.
[Excel IP Functions.bas](VBA/Excel%20IP%20Functions.bas)
[Excel IP Functions.cls](VBA/Excel%20IP%20Functions.cls)

Easy Installation:
Download the blank workbook and start using the functions:
[Excel IP Functions.cls](VBA/Excel%20IP%20Functions%20Blank%20Workbook.xlsm)
