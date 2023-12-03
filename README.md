# EasyFileManager

With this application, it is possible to manage your computer files in batch, based on their metadata. Files can be modified by editing metadata, but can also be renamed and moved (or copied). The batch process can be extended by finding and removing duplicates and/or deleting empty folders. Settings are saved when changes are made, but can also be exported and imported. There are additional options, such as preserving dates of edited and/or moved (or copied) files. Files can be opened directly or viewed in *Windows Explorer*. Images and videos can be previewed, and images can be rotated on the fly.

There are six languages available to choose from: *German*, *English*, *Spanish*, *French*, *Italian* and *Dutch*. Detailed information on how to use *EasyFileManager* can be found in the [wiki](https://github.com/rafcolson/EasyFileManager/wiki).

![EFM](https://github.com/rafcolson/EasyFileManager/assets/10002909/65289d0a-4a7c-4fc6-989b-fa48eb6fccb1)

License
-------
*EasyFileManager* is free software, but not license-free, mainly because it uses software and services under the terms of several licenses:
- [Microsoft Visual Studio Community](https://visualstudio.microsoft.com/vs/community/) under the [Microsoft Software License for Visual Studio Community](https://visualstudio.microsoft.com/license-terms/vs2022-ga-community) for individuals
- [Windows-API-Code-Pack-1.1](https://github.com/aybe/Windows-API-Code-Pack-1.1) under the [Microsoft Windows API Code Pack License](https://web.archive.org/web/20130717101016/http://archive.msdn.microsoft.com/WindowsAPICodePack/Project/License.aspx)
- [ExifTool](https://exiftool.org) under the [Perl Artistic License](https://dev.perl.org/licenses/artistic.html)
- [OpenStreetMap](https://www.openstreetmap.org) under the [Open Data Commons Open Database License (ODbL)](https://opendatacommons.org/licenses/odbl/) by the [OpenStreetMap Foundation (OSMF))](https://osmfoundation.org).

As for code created by myself (or other contributors), *EasyFileManager* may be redistributed and/or modified under the [BSD 3-Clause License](https://opensource.org/licenses/BSD-3-Clause), also ensuring none of its contributors are liable for any damages which may be caused while using the software, when either customizing, renaming, moving or deleting computer files.

The licenses which users must agree to in order to install this software can be viewed in [LICENSE](https://github.com/rafcolson/EasyFileManager/blob/main/LICENSE).

Development
-----------
Since I was tired of manually organizing photos and videos for family, I decided to start this project to make it easier to go through terrabytes of data. As I gradually learned about existing software and services, it quickly evolved into an easy, yet somewhat advanced tool for managing computer files.

Credit to *Microsoft* for allowing developers to create for free, with their *Visual Studio Community*. Also kudos to the contributors to the *Microsoft Windows API Code Pack* project for their updates, making it (more) compatible with newer *.NET* packages. Credit to the creator of *ExifTool* for allowing other developers to freely attain access to embedded metadata of videos. Especially praise to the *OpenStreetMap Foundation* for sharing their resources of the geographically mapped world, with the world. Without any of those parties *EasyFileManager* would not have been possible, not only in terms of time, but also in general, as I am not a professional programmer.

Dependencies
------------
With the intension of reusing large portions of the code for other software applications, this *Microsoft Visual Studio project* is referencing one of my other projects: *WinFormsLib*, which is part of the rather custom-made .NET library [DotNetHelper](https://github.com/rafcolson/DotNetHelper).

Requirements
------------
*EasyFileManager* will run on Windows 7 to 11.

Release notes
-------------
See [CHANGELOG](https://github.com/rafcolson/EasyFileManager/blob/main/CHANGELOG) for the latest changes.

Download and Install
--------------------
Releases are published [here](https://github.com/rafcolson/EasyFileManager/releases). Click on *Assets*, download and run *EasyFileManagerSetup-x64.msi*.
