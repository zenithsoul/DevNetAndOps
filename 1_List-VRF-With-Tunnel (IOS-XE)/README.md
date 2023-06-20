# List VRF With Tunnel

Here's an example of coding for network operations that gathers information for VRF and matches only tunnel interfaces, then generates an Excel file.

## Requirement

> pip install -r requirements.txt
>
> Support IOS-XE Gibraltar 16.11.1 or later (I test on IOS-XE 17.3.1)
>
> Enabling RESTCONF on IOS-XE
> https://developer.cisco.com/docs/ios-xe/#!enabling-restconf-on-ios-xe
>
> and check the device must run "RESTCONF" or add configure
>
> Router(config)#resetconf
>

## Before run the script

Change the IP Address
> IP_Device = "127.0.0.1"

Change the Username and Password
> username = "example"
> password = "example"

and The name of an excel file.
> output_file = "example.xlsx"
