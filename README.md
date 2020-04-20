# VBS-Toolkit
>VBS Toolkit with useful ready to use functions

![screeshot](https://lh3.googleusercontent.com/3L-y_gu-tj1RbtWj-SnvG2TGhMqC_1XkqCTC4ZsPvov0OMbB6qH5oo1o5xTIyf924hwcIQ)
## Available Functions
### Write Logs / Text file
- WriteTextFile

### Run Commands
- RunLocalCommand
- RunRemoteCommand

### Active Directory
- FindADObject
- CheckADPathName

### Windows Management
- GetEnvPath
- GetServiceStatus
- GetEvents

### WMI
- CheckWMIAccess
- WMIDateStringToDate
- GetWMICollectionItem

### Registry
- RegReadRemoteValue
- RegKeyExsists
- RegValueExsists
- RegWriteRemoteValue
- RegReadLocalValue
- RegWriteLocalValue

### Email
- SendMail

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
