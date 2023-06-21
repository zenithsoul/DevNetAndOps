# List Tunnel Interface and check VRF on the Tunnel Interface

This guide demonstrates an example of network operations code that gathers Tunnel Interface information and check VRF on the Tunnel Interface. The program then generates an Excel file with this information.

## Python Package Requirements

To install the required Python packages, run the following command:

> pip install -r requirements.txt

## Device Requirements

Your device must meet the following criteria:

- **Support IOS-XE Denali 16.3.1 or later (Tested on IOS-XE Amsterdam 17.3.1)**
- Enable RESTCONF on IOS-XE. Learn how to enable it here. <https://developer.cisco.com/docs/ios-xe/#!enabling-restconf-on-ios-xe>
- Ensure that the device runs "RESTCONF" or add the following configuration:

> Router(config)#resetconf

## Before Running the Script

Update the following variables in the script as needed:

Change the IP Address:
> IP_Device = "127.0.0.1"

Change the Username:
> username = "example"

Change the Password:
> password = "example"

Change the name of the output Excel file:
> output_file = "example.xlsx"
