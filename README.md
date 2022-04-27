# ie4w11

Hello all!  

By default, Windows 11 will redirect all Internet Explorer instances to the Edge browser unless you add the URLs to "Internet Explorer mode pages" but only for 30 days, and you have to extend the expiry day every 30 days which is very inconvenient.  

I am working in a hotel as an IT. A core application must be run on Internet Explorer with JRE installed but the Edge browser cannot handle. Thus, I made this vbscript for running native Internet Explorer on Windows 11 which also supports URL as a parameter.  

While starting ie4w11.vbs, it will terminate all hidden iexplore.exe and ielowutil.exe before creating a new instance of iexplore.exe.  

Note:  
The default script engine for Windows is wscript.exe. If you have changed the default engine to cscript.exe, a black console window will be pop-up. You can suppress this black console window by creating a shortcut and type "wscript.exe ie4w11.vbs" as the target.  

## Disclaimer
This software is provided as-is with no warranty. I am not an expert and I am not liable for any accidental damage to your hardware, systems or files. The software has access to disks and partitions, and it can erase a disk if something goes wrong.  

You are free to use the software, or edit the script for your needs, in both non-commercial and commercial use.  

## ie4w11.vbs  
Main vbscript for running the native Internet Explorer on Windows 11

**Syntax**  
**ie4w11.vbs**  
Start Internet Explorer with a homepage. It will look for the homepage from the following locations and sequence.  
1. HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Internet Explorer\Main\Start Page  
2. HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Start Page  
3. HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Edge\HomepageLocation  
4. HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Edge\Recommended\HomepageLocation  
5. HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\MicrosoftEdge\Internet Settings\HomeButtonURL  

If no homepage is found, the homepage will be default to about:blank  

**ie4w11.vbs https://self-url**  
Start Internet Explorer with https://self-url.  

## ie4w11.exe.zip  
This zip file contains  
- ie4w11.exe
- ie4w11.dll
- ie4w11.exe.manifest  

**Usage**  
Unzip the zip file and put all these three files together in the same folder.  

## ie4w11.exe  
ie4w11.exe is complied by VbsEdit from ie4w11.vbs which does the same thing as ie4w11.vbs.  

**Syntax**  
Same as ie4w11.vbs.

## Dependencies  
Microsoft.NET Framework v2.0

## Tested Windows version
- Windows 10
- Windows 11

