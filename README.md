# VBS-Toolkit
VBS Toolkit with useful ready to use functions

![screeshot](https://lh3.googleusercontent.com/3L-y_gu-tj1RbtWj-SnvG2TGhMqC_1XkqCTC4ZsPvov0OMbB6qH5oo1o5xTIyf924hwcIQ)
## Available Functions
### Write Logs / Text file
- WriteTextFile
	- Write appended data to text file. Creates the File if it does not exists.

### Run Commands
- RunLocalCommand
	- Runs a command in the Windows Shell and Returns its Output
- RunRemoteCommand
	- Runs a Remote Command Using WMI on a Remote Computer

### Active Directory
- FindADObject
	- Finds Objects in the Active Directory by Object Class and Search by Category
- CheckADPathName
	- Checks the AD Path Name for Special Characters and Replaces them with the Proper Escape Characters

### Windows Management
- GetEnvPath
	- Converts String Environment Constants to the Environment Path Requested
- GetServiceStatus
	- Checks the Service State and returns it
- GetEvents
	- Return an Array of Events Between Dates. This Function Uses LogParser as the Event Log Collector.

### WMI
- CheckWMIAccess
	- Checks WMI Access on a Computer
- WMIDateStringToDate
	- Converts the WMI Date time Format to Normal Date Time
- GetWMICollectionItem
	- Connects to the WMI Service and Runs a Query

### Registry
- RegReadRemoteValue
	- Read a Value from a Remote Registry Path
- RegKeyExsists
	- Check if a Key from a Remote Registry Path exists
- RegValueExsists
	- Check if a Value from a Remote Registry Path exists
- RegWriteRemoteValue
	- Write a Value to a Remote Registry Path
- RegReadLocalValue
	- Read a Value from a Local Registry Path
- RegWriteLocalValue
	- Write a Value to a Local Registry Path

### Email
- SendMail
	- Sends an Email from the Users Outlook to a Destination Address

## How to use
### Create an instance
You can call the toolkit from other VBS scripts like any other CPM object.
If you [register the toolkit on your host](Register-the-Toolkit-on-your-machine), you can call it using for example:
```vbs
Set oComponent = CreateObject("Toolkit.AssafM.wsc")
```
If you choose not to register the toolkit, you can use the following example:
```vbs
Set oComponent = GetObject("script:C:\Scripts\Toolkits\Toolkit.AssafM.wsc")
```
> Note: You can use it just the same from C++ or HTML

### Register the Toolkit on your machine
In Windows Explorer, right-click the script component (.wsc) file, and then choose Register.
–or–
Use the new version of Regsvr32.exe that ships with the script component package, and use this command:
regsvr32 ToolKit.AssafM.WSC

## License
This repository is licensed under Apache License 2.0 - see [`LICENSE`](LICENSE) for more details.
=======
 VBS Toolkit with useful ready to use functions
>>>>>>> 88812c7f575af58f7ac04bb497055b5eaa0975d5
