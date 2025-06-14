Simple mbox viewer
=======

A simple but powerfull viewer to view mbox files such as Thunderbird Archives, Google mail archives or simple Eml files.
[![Download Windows MBox Viewer](https://img.shields.io/sourceforge/dm/mbox-viewer.svg)](https://sourceforge.net/projects/mbox-viewer/files/latest/download)

Supported OS
------------

MBox Viewer is built on Windows WIN7 platform.
Version 1.0.3.47 is the last release built on Windows XP platform.

Features
--------

* large file support > 4Gb;
* fast parsing of mbox;
* quick access to attachments;
* preview picture attachments;
* export of all attachments;
* export of single mail in Eml;
* export of all mails to separate Eml files;
* print all or multiple selected mails to CSV or Text or HTML or PDF file;
* print single mail to Text or HTML or PDF file or to send to PDF printer;
* group all related mails as conversations in the summary window;
* print related mails/mail group to Text or HTML or PDF or CSV files or send to PDF printer;
* open a single or multiple selected mails or group of related mails in external browser;
* sort emails by date, from, to, subject and conversation columns;
* find user defined text in the message and highlight all occurrences;
* search mails by user defined date range and text in the subject, sender, recepient, cc, bcc, message text, attachments text and attachment names;
* search for mails that didn't match the search criteria specified by a user;
* support for header fields and body text encoded with any Windows supported character sets/page codes;
* ability to compose mail list from single or multiple search results and archiving if desired;
* ability to set Message Window position to Bottom, Right or Left;
* ability to merge/concatenate multiple archive files and remove duplicate mails;
* ability to customize the background color of key display panes;
* ability to customize the HTML/PDF mail header output;
* view raw mail header;
* forward mails directly from MBox Viewer;
* support for Gmail Labels;
* support for Thunderbird exported mail files;
* support for multiple languages;
* support for file based data configuration in addition to the default Registry based;
* ability to configure font size of GUI objects;

License
-------
Mbox viewer is a free software. Most of it is licensed under AGPLv3.
Small parts are licenced under CPOL 1.02.
MailKit open source library linked to connect to SMTP servers is licensed under MIT.
See the [LICENSE.txt](LICENSE.txt) file included in the distribution for further details.

NOTE: The source code and executable were released under GPLv2 on Sourceforge since the project inception as shown on the Summary page.
There was no explicit license information for the project on the Github. However, without a license, the default copyright laws apply, 
meaning that authors retain all rights to the source code and no one may reproduce, distribute, or create derivative works from the work.
The [LICENSE.txt](LICENSE.txt)  file was added in v1.0.3.4 to make the licensing explicit and to synchronize the License published on Sourceforge and Github.


Documents
---------

[UserGuide.pdf](UserGuide.pdf) provides comprehensive information how to install MBox Viewer, start and use supported features.

Videos
------

Refer to [Introduction to MBox Viewer](https://www.youtube.com/watch?v=qrjjR9Bvz8k) video (somewhat outdated) for introduction.


Installation
------------

MBox Viewer executable is released as mbox-viewer.exe-v1.0.3.XX.zip file. Please install zip file in selected folder and unzip the zip file.
This wil create mbox-viewer.exe-v1.0.3.XX subfolder under the selected folder. 
The mbox-viewer.exe-v1.0.3.XX subfolder will contain two executable files mboxview.exe or mboxview64.exe.

Configuration
-------------

MBox Viewer supports Windows Registry based and the file based configuration.

By default, Windows Registry is used to store configuration data.
During startup, the MBox Viewer will check whether the MBoxViewer.config file exists and is writeable in:

1. the Config subfolder under the MBox Viewer software installation folder  or
2. in the UMBoxViewer\Config subfolder under the  user specific folder created by Windows system 
        example : C:\Users\UserName\AppData\Local\UMBoxViewer\Config

The config file format is similar to the format of ".reg" registry file

```
[UMBoxViewer\LastSelection]
"parameter"="value"
```

White spaces are not allowed in the front of each line and around "=" character.
All parameter values are encoded as strings and converted by MBox Viewer to numbers or other data types when needed.

MBoxViewer.config file must be encoded as UTF16LE BOM file

MBoxViewer.config.sample file is included in the software package under the Config folder.
In order to enable MBox Viewer to use the file based configuration, 
user needs to rename this file to MBoxViewer.config file or copy the sample file
to C:\Users\UserName\AppData\Local\UMBoxViewer\Config folder and rename.

Multiple Languages
------------------

MBox Viewer supports the multiple languages:

English, Italian, Spanish, German, French, Portuguese, Portuguese-Brazil, Romanian, Polish, Japanese and Chinese-Simplified

GUI strings and documentation were translated from English to target languages using free Google Translation service and were not reviewed by humans yet. 
Free Google Translations are not domain specific and most likely they will need to be updated by humans.

By default English language is enabled. User can select different language during runtime by selecting 'Language->Select Language' option on the main menu.


Running
--------

Open created mbox-viewer.exe-v1.0.3.XX subfolder. Double left click on the mboxview.exe or mboxview64.exe file to start MBox Viewer. 
You will be asked to configure folder for all temporary files created by MBox Viewer.

On the main top tool bar, select the `File` option to open the drop menu and then select the `Select folder...` option. 
Browse to the folder containing one or more mbox and/or eml mail archive files and select it. All valid mail archive files will appear in the Mail Archive Tree window.

Left click on one of the archive files to load all mails within that archive.  Progress bar will appear and automatically close after the selected archive is fully processed. 
Mail header information of each email will appear in the Summary window. 
Note that parsing of very large archive file may take some time since the mail archives are text files and every character has to be examine one by one.  
However, subsequent loading of mails is done from the index file created by the mboxview during the initial parsing of the archive file and is much faster.
The created index file contains content meta data of each mail in the archive file. i.e. the mail header information and the position of each mail within the mail file for quick access to the mail message/body. 
The index files have the .mboxview extension.
	   
Left Click on one of the mails in the Summary window to show the Message/Body of that email in the Message window.
The mail retrieval state, total number of mails in the archive and the position of the selected mail within the archive is shown on the status bar.
	   
Refer to [Introduction to MBox Viewer](https://www.youtube.com/watch?v=qrjjR9Bvz8k) video for basic information on how to run MBox Viewer.

Refer to provided [UserGuide.pdf](UserGuide.pdf)  how to run MBox Viewer from command line.

Font Size
----------

User has an option to increase the font size of GUI objects if desired. On the main top menu tool bar, select the `File` option to open the drop menu and then select the `Font Config` option. 

Changes
-------

See [CHANGE_LOG.md](CHANGE_LOG.md) file.

Building
--------

MBox Viewer project executables are built under VS 2022. Project contains two solution files:

**mboxview.sln** - this will build Release and Debug versions of mboxview.exe and mboxview64.exe executables

**ForwardEmlFile/ForwardEmlFile.sln** -  this will build  Release and Debug versions of ForwardEmlFile.exe


