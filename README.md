# EasyFileManager

*EasyFileManager* is an open-source project developed and maintained by Raf Colson under the name RacoSoft.

With this application, it is possible to manage your computer files in batch, based on their metadata. Files can be modified by editing metadata, but can also be renamed and moved (or copied). The batch process can be extended by finding and removing duplicates and/or deleting empty folders. Settings are saved when changes are made, but can also be exported and imported. There are additional options, such as preserving dates of edited and/or moved (or copied) files. Files can be opened directly or viewed in *Windows Explorer*. Images and videos can be previewed, and images can be rotated on the fly.

There are six languages available to choose from: *German*, *English*, *Spanish*, *French*, *Italian* and *Dutch*. Detailed information on how to use *EasyFileManager* can be found in the [wiki](https://github.com/rafcolson/EasyFileManager/wiki).

![EFM](https://github.com/rafcolson/EasyFileManager/assets/10002909/65289d0a-4a7c-4fc6-989b-fa48eb6fccb1)

License
-------
*EasyFileManager* is free of charge and open-source software. Its source code and the third-party software, services, and data it uses remain subject to their respective licenses:
- [ExifTool](https://exiftool.org) under the [Perl Artistic License](https://dev.perl.org/licenses/artistic.html)
- [OpenStreetMap](https://www.openstreetmap.org) data under the [Open Data Commons Open Database License (ODbL)](https://opendatacommons.org/licenses/odbl/)

Code developed specifically for *EasyFileManager* by Raf Colson and other contributors is licensed under the [BSD 3-Clause License](https://opensource.org/licenses/BSD-3-Clause), which includes a disclaimer of liability.

The applicable license terms and notices for this software can be viewed in [LICENSE](https://github.com/rafcolson/EasyFileManager/blob/main/LICENSE).

Credits
-------
Credit to *Microsoft* for allowing developers to create software for free with *Visual Studio Community*. Also credit to the creator of *ExifTool* for giving other developers free access to embedded metadata in videos. Special thanks to the *OpenStreetMap Foundation* for sharing their resources of the geographically mapped world with the world. Without any of those parties, *EasyFileManager* would not have been possible, not only in terms of time, but also in general, as I am not a professional programmer.

Development
-----------
Since I was tired of manually organizing photos and videos for family, I decided to start this project to make it easier to go through terabytes of data. As I gradually learned about existing software and services, it quickly evolved into an easy, yet somewhat advanced tool for managing computer files.

The project can be built with [Microsoft Visual Studio Community](https://visualstudio.microsoft.com/vs/community/), subject to the terms of [Microsoft Software License for Visual Studio Community](https://visualstudio.microsoft.com/license-terms/vs2022-ga-community) for individuals. The Windows Forms Designer layout was created at 1920 × 1080 with Windows display scaling set to 100%. Use these settings for the most accurate design-time representation.

Dependencies
------------
With the intention of reusing large portions of the code for other software applications, this *Microsoft Visual Studio project* references one of my other projects: *WinFormsLib*, which is part of the rather custom-made .NET library [DotNetHelper](https://github.com/rafcolson/DotNetHelper).

Requirements
------------
*EasyFileManager* will run on Windows 7 to 11.

Release notes
-------------
See [CHANGELOG](https://github.com/rafcolson/EasyFileManager/blob/main/CHANGELOG) for the latest changes.

Download and Install
--------------------
Releases are published [here](https://github.com/rafcolson/EasyFileManager/releases). Click on *Assets*, download and run *EasyFileManagerSetup-x64.msi*.
