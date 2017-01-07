# netsh to clipboard
A simple MS Excel macro that copies to the clipboard 'netsh' commands (Windows) that change network settings for Local Area Connection, based on selected cells (which ideally contain IP addresses). The copied content can then be pasted to an elevated command prompt to execute it. You can set a keyboard shorcut for this macro, e.g. CTRL+SHIFT+I, for faster results.

Just in case you need to check a large set of static IP addresses and DNS servers for internet connectivity. :)

## Usage
Select the cells (IP addresses) corresponding to the IP address, subnet mask, default gateway, primary DNS, and secondary DNS, and hit the keyboard shorcut to copy the command to the clipboard. Example of generated command:

```
netsh interface ip set address "Local Area Connection" static ip_here subnet_mask_here gateway_here 1 &&
netsh interface ip set dns name="Local Area Connection" source=static addr=primary_dns_here register=primary &&
netsh interface ip add dns name="Local Area Connection" addr=secondary_dns_here index=2
```

Copy the result to an elevated command prompt to set the network settings.
