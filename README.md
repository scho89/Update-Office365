### [Update-Office365.ps1](https://github.com/chosangho89/Update-Office365/raw/master/Update-Office365.ps1) : Script for Update or Rollback Office 365 Client build
##### You can also install this script using cmdlet "Install-Script Update-Office365" <http://scho.kr/update365psg>  
Former https://aka.ms/update365 

![Update-Office365](/info.png)

Description
==============
-This sciprt helps you to update or rollback your Office 365 client.   
-You can choose your build number or channel.   
-This sciprt change your update channel for Office 365 client.   
-You may need to change execution policy of your system. (Run : Set-ExecutionPolicy -ExecutionPolicy RemoteSigned #more detail : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-6)   
-You may need to unblock this script. (User Unblock-File cmdlet or right click - Properties - Unblock)   
-You need to run Internet Explorer if you get a blank drop down list of Version field. (Initial setting for IE is required)   
  
Details
===========
MD5 : 2E86D9F60263270622E9A6BDB13CDF10   
SHA1 : 00E44D64C46F7D278AF7DDA9ED2CD218B22F9929   
SHA256 : B79449918AE69E09A4606734FE7C729F7AB6FC5398B8FE05C2788CFC862F291E   
File size : 14.4KB (14,813 bytes)   
File type : PS1   

Requirements
============
Office 365 Business, Office 365 ProPlus, Office 365 Personal   

Release note
==============
2020-06-26 : Channel name changed.   
2019-04-19 : Verification procedure for installed Office 365 client  is added.   
2019-02-26 : Bud fixed / SSL/TLS issue.   
2018-10-16 : Insider, Monthly Channel (Targeted), DevMain channel added.   
2018-10-02 : Bud fixed / Original content URL is changed.   
2018-06-25 : Bug fixed / Deferred channel entry is displayed correctly.   
2018-06-08 : Updated (Source content is updated!. Modified regular expression) / Add check box for disabling update.   
2018-03-06 : Bug fixed / Now, you can change channel again!   
2017-09-26 : Bug fixed / Insider channel removed   
2017-07-26 : GUI is added   
2017-07-25 : This script can modify registry value for changing update channel.   
2017-07-24 : Insider channel (First release for current channel) is added.   
