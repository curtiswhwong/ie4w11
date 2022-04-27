# ie4w11

Hello all!  

By default, Windows 11 will redirect all Internet Explorer instances to the Edge browser unless you add the URLs to "Internet Explorer mode pages" but only for 30 days, and you have to extend the expiry every 30 days which is very inconvenient.  

I am working in a hotel as an IT. A core application must be run on Internet Explorer with JRE installed but the Edge browser cannot handle. Thus, I made this vbscript for running native Internet Explorer on Windows 11 which also supports URL as a parameter.

## ie4w11.vbs  
Main vbscript for running the native Internet Explorer on Windows 11

Syntax:  
**ie4w11.vbs**  
Start Internet Explorer with a preset homepage. It will look for the homepage from the following locations.  
HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Internet Explorer\Main\Start Page  
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\Start Page  
HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Edge\HomepageLocation  
HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Edge\Recommended\HomepageLocation  
HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\MicrosoftEdge\Internet Settings\HomeButtonURL  
If no homepage is found, the homepage will be default to about:blank  

**ie4w11.vbs https://self-url**  
Start Internet Explorer with https://self-url.
