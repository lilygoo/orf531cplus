Project xlwtest
===============
Michael G Sotiropoulos, 24-Jun-2017

This project depends on the `xlw` third party library.  
Make sure that you have unzipped the provided `xlw-5.0.2f0.zip` file inside your `Develop` folder.


How to adjust the include and library paths
-------------------------------------------

* Open the `scratch-vs15.sln` solution file in Visual Studio 2017.
* Open the `xlwtest` project property pages.
* Set "Configuration" to "All Configurations" and go to Configuration Properties -> C/C++
* Set "Additional Include Directories" to "\<PathToDevelopFolder\>/xlw-5.0.2f0/xlw/include"

* Go to Configuration Properties -> Linker
* Set "Additional Library Directories" to "\<PathToDevelopFolder\>/xlw-5.0.2f0/xlw/lib/$(PlatformTarget)"

The above information is saved in the project file `xlwtest-vs15.vcxproj`.

How to configure debugging with Excel
-------------------------------------

* From Visual Studio open the `xlwtest` project property pages.
* Set "Configuration" to "All Configurations" and go to Configuration Properties -> Debugging
* Set "Command" to the path to your `excel.exe` executable.  
	The Excel 2013 64-bit executable is in  
	`C:/Program Files/Microsoft Office/Office15/excel.exe`  
	The Excel 2013 32-bit executable is in  
	`C:/Program Files (86)/Microsoft Office/Office15/excel.exe`  
	If you have Excel 2016, change "Office15" to "Office16" in the above path.  
	Depending on the installer, sometimes the path to `excel.exe` may have a `root` subdirectory, like  
	`C:/Program Files/Microsoft Office/root/Office15/excel.exe`

	For 64-bit Excel, choose the `x64` platform version of the `xlwtest` project. 
	For 32-bit Excel, choose the `x86` platform version of the `xlwtest` project. 	
	
* Set "Command Arguments" to "$(TargetPath)" and click OK.

The above information is *not* saved in the `xlwtest-vs15.vcxproj` but in the user-specific XML file `xlwtest-vs15.vcxproj.user`.

When you hit F5, Visual Studio may pop-up a dialog window saying that there is no Debugging Information for Excel.  
Check the "Don't show this dialog again" box and click Yes.

If Excel pops-up a security notice and you want to suppress it for ever, you need to reduce Excel's security settings.  
In Excel 2013 and 2016 this is done from  
File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings  
Click on "Enable all macros".
