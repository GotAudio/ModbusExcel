ModbusExcel v 0.1 Alpha (Aug 13th, 2015)
===

RTD server that polls Modbus TCP registers from Excel

### Getting Started
- At this time, only sources are provided, no installer or binaries.
- Use Git to clone branch or download and unzip .ZIP file to your hard drive.
- The first time you build the ModbusExcel project you must launch Visual Studio as Administrator in order to register RTD COM compoent
- Open ModbusExcel.sln with Visual Studio. Tested on Win7 VS 2013 and Win10 VS 2015.
- Right click solution name in VS Solution Explorer and click Build. NuGet packages will be automatically downloaded.
- Launch ModbusSim.exe
- Click "StartListening"
- Open Documents\ModbusExcel.ReadCoils.FormatRules.xlsm with Excel
- Click on the Red circle in cell A1 to start polling. Circle should turn green and values should appear. Click on Green circle to stop polling.
- Try some of the other spreadsheets in the Documents folder. 
	Note: Closing and opening different workbooks, closing workbooks without stopping polling, or opening multiple polling workbooks simultaneously can have unpredictable results.  It is best to stop polling and close the current workbook before opening a new one.)
- There are various configuration settings available in columns C and D. I will document their use soon.
- Installation and user guide is here (work in progress): https://github.com/GotAudio/ModbusExcel/blob/master/Documents/ModbusExcel%20User%20Guide.docx

This solution contains four projects;

### ModbusExcel	- RTD Server for Excel. (See ModbusExcelGuide.xls for more details)

### ModbusSim	- Load test simulator to test performance and connectivity.
- ModbusSim is not a fully functional ModBus server. The responses it sends are valid for the register datatype, with some customization for Omni Flow Computer registers (eg. valid ASCII and numeric Dates) It ignores UnitID so they all retuen the same result. It does support Omni RDA (Raw Data Archive 700-7nn) Method 1 request of Single read Address:Index. Method 2 (write first) is not supported. It never returns an error status.  What it does do is accept tens of thousands of simultaneous async connections and async reads. You can configure the Excel spreadsheet to use the RTD service to make a new connection for EVERY register, rather than the default of reusing one connections for all requests for a device. and then continuously read thousands of registers as fast as possible. On my PC, the RTD service is able to establish 3000 connections in 10 seconds, and make and receive 110,000 poll responses in 60 seconds using no more than 3 threads, minimal memory, and less than 10% CPU. 
- Also for what it's worth, even though the screen looks like a grid of checkbox controls, it is actually just a bitmap image of a grid. As connections are made, and read requests are received, 3-state checkboxes appear to change. Those checkboxes are just bitmaps drawn on the screen at the appropriate locations.  This was the only way to manage to toggle thousands of checkboxes per-second, as opposed to the tens of checkbox changes per-second MSControls support.  Even now, those checkboxes add a 30% overhead to the simulator. I am working on an alternate scheme. You can uncheck "Show Events" to disable realtime connection and data indicators.

### ModbusExcel.Tests	- Incomplete (skeleton) tests. TBD: Tests need to be written and converted from MSTest to Nunit.

### ModbusExcelConsole	- An RTD test console. ModbusSim.exe must be running for this to work.
- Modified from; https://github.com/AJTSheppard/Andrew-Sheppard/blob/master/MyRTD/MyRTDConsole/Program.cs (Can probably be deleted after tests are complete.)


### ModbusExcel uses the following components;

- ModbusTCP Class	- http://www.codeproject.com/Tips/16260/Modbus-TCP-class (Modified "Nov-6th-2006" ModbusTCP.cs version included)
- NitoAsync		- https://nitoasync.codeplex.com/ (NuGet package v1.4. Included upgraded .Net 4.0 projects with this solution.)
- Nlog			- http://nlog-project.org/ (NuGet package v4.0.1)
- Nunit			- http://nunit.org/ (NuGet package v2.6.4)


### Notes: 
- See ModbusExcel\NLog.config to change output log file name and location (if logging is enabled. See flags in Logging worksheet).
- You can find another modbus simulator here http://www.plcsimulator.org/. This https://sourceforge.net/projects/modrssim2 may be a later fork. I have included a spreadsheet, (ModbusExcel.PLCSimulator.xlsm), already setup to poll this simulator.  A word of caution though, this simulator will crash if you establish too many connections, poll it too quickly, or if you make many async requests. It also seems to drop a lot of connections. But it has more features and offers more control over responses and data than ModbusSim.


Ken Selvia (gotaudio at gee mail dot com)

```
TODO:	Complete Documents\ModbusExcelGuide.docx and generate and publish .pdf version.
		Include binaries and/or installer
		Document MSMQ integration and SQL CLR
		Include sql-msmq-modbus.sql		
		Implement even faster graphics drawing in ModbusSim. Without drawing, throughput increases from 110K to 180K reads per minute. 
		Convert ModbusSim to WPF Metro theme
		Write tests for ModbusSim
```