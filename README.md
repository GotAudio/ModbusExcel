# ModbusExcel

RTD server that polls Modbus TCP registers from Excel

This solution contains four projects;

1.	ModbusExcel	- RTD Server for Excel. (See ModbusExcelGuide.xls for more details)

2.	ModbusSim	- Load test simulator to test performance and connectivity.

	Note:	
	ModbusSim is not a fully functional ModBus server. The responses it sends are valid for the register datatype, with some customization 
			for Omni Flow Computer registers (eg. valid ASCII and numeric Dates) It ignores UnitID so they all retuen the same result.
			It does support Omni RDA (Raw Data Archive 700-7nn) Method 1 request of Single read Address:Index. Method 2 (write first) is not supported.
			It never returns an error status.  What it does do is accept tens of thousands of simultaneous async connections and async reads.  
			You can configure the Excel spreadsheet to use the RTD service to make a new connection for EVERY register, rather than the default of
			reusing one connections for all requests for a device. and then continuously read thousands of registers as fast as possible.
			On my PC, the RTD service is able to establish 3000 connections in 10 seconds, and make and receive 110,000 poll responses in
			60 seconds using no more than 3 threads, minimal memory, and less than 10% CPU. 

			Also for what it's worth, even though the screen looks like a grid of checkbox controls, it is actually just a bitmap
			image of a grid. As connections are made, and read requests are received, 3-state checkboxes appear to change.
			Those checkboxes are just bitmaps drawn on the screen at the appropriate locations.  This was the only way to manage to
			toggle thousands of checkboxes per-second, as opposed to the tens of checkbox changes per-second MSControls support.

			
3.	ModbusExcel.Tests	- Incomplete (skeleton) tests. TBD: Tests need to be written and converted from MSTest to Nunit.

4.	ModbusExcelConsole	- An RTD test console.

	Note:	Modified from; https://github.com/AJTSheppard/Andrew-Sheppard/blob/master/MyRTD/MyRTDConsole/Program.cs

			Can probably be deleted after tests are complete.


Uses the following add-ins or components;

1.  ModbusTCP Class	- http://www.codeproject.com/Tips/16260/Modbus-TCP-class (Modified "Nov-6th-2006" ModbusTCP.cs version included)

2.	NitoAsync		- https://nitoasync.codeplex.com/ (NuGet package v1.4. Update 8-10-15: Included upgraded .Net 4.0 projects in this solution.)

3.	Nlog			- http://nlog-project.org/ (NuGet package v4.0.1)

4.	Nunit			- http://nunit.org/ (NuGet package v2.6.4)


Notes: 
	Opening more than one spreadsheet that uses "ModbusExcel.RTD" will have unpredictable results.

	See ModbusExcel\NLog.config to change output log file name and location.

	You can find another modbus simulator here http://www.plcsimulator.org/. This https://sourceforge.net/projects/modrssim2 may be a later fork. I have included a spreadsheet, (ModbusExcel.PLCSimulator.xlsm), already setup to poll this simulator.  A word of caution though, this simulator will crash if you establish too many connections, poll it too quickly, or if you make many async requests. It also seems to drop a lot of connections. But this one has more features and offers more control over responses and data.


Ken Selvia (gotaudio at gee mail dot com) August 7th, 2015

	TODO:	Document MSMQ integration and SQL CLR
			Include sql-msmq-modbus.sql		
			Convert ModbusSim to WPF Metro theme
			Write tests for ModbusSim
			Complete Documents\ModbusExcelGuide.xlsx
			Implement even faster graphics drawing. Without drawing, throughput increases from 110K to 180K reads per minute. 
			
		

