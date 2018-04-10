Simple mbox viewer
=======

A simple viewer to view mbox files such as Thunderbird Archives, Google mail archives or simple Eml files.
[![Download Windows MBox Viewer](https://img.shields.io/sourceforge/dm/mbox-viewer.svg)](https://sourceforge.net/projects/mbox-viewer/files/latest/download)

Features
-----

* large file support > 4Gb;
* fast parsing of mbox;
* quick access to attachments;
* export of single mail in Eml;
* search by date, subject, sender or raw email; 
* sort date, from, to and subject columns;
* support for ISO-8859-[1-9] and utf8 encoded header fields and body text;
* support for searching emails sorted by header fields (except for searching of entire RAW message);

Changes
---

v 1.0.2.1

 - To improve mbxoview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - updated User Manual to describe new and updated features;
 - enhanced search criteria to include from, to, subject, message and text attachments elements;
 - deprecated searching of RAW message;
 - eliminated false negative searches;

v 1.0.1.13

 - updated User Manual to describe new and updated features;
 - fixed header fields presentation in the header of Message Window;
 - changed the layout of the header in the Message Window;
 - added new gobal options settable from GUI;
 - fixed occasional truncation of the subject field;
 - fixed occasional blank Message Window during traversal and searching;

v 1.0.1.12

 - added User Manual;
 - added delayed search progress bar to view search progress and to allow to cancel the search in progress;
 - added -PROGRESS_BAR_DELAY=seconds command line option to control the start the progress bar. The value of -1 disables the bar;
 - addedd ability to set/override EXPORT_EML and PROGRESS_BAR_DELAY from GUI;
 - fixed small issue with searching of RAW messages;
 - note: -MAIL_FILE=filename new option was stated incorrectly as -EMAIL_FILE=filename in the v1.0.1.11 READMEs;

v 1.0.1.11

 - added email size column;
 - added automatic column resize upon window resize;
 - added -EMAIL_FILE=filename command line option to open specific file archive;
 - added -EXPORT_EML=y|n command line option to suppress export of eml file to improve performance;
 
v 1.0.1.10

 - enhanced support for ISO-8859-[1-9] encoded text content;
 - added support for searching emails sorted by header fields (except of searching of entire RAW message);
 
v 1.0.1.9

 - added support for ISO-8859-[1-9] encoded text content
 
v 1.0.1.8

 - added support for ISO-8859-[1-9] encoded from, to and subject fields

v 1.0.1.7

 - more fixes by Zbigniew Minciel (change for test)

v 1.0.1.5

 - added sorting thanks to Zbigniew Minciel

v 1.0.1.4

 - fixed a bug where program would hang
 - fixed save state of emails

v 1.0.1.3

 - can specify folder from command line with -FOLDER= option

v 1.0.1.2

 - added donation link in about box (displayed on each version change)
 - fixed character encoding problems
 - fixed correct display of progress bar
 - made parsing faster
 - added View EML in View menu

v 1.0.1.0

 - changed version number because of a problem with sourceforge

v 1.0.0.10

 - minor update to fix the display of mime encoded subject

v 1.0.0.9

 - date formatting in the Options command in the File menu;
 - display of images in html emails;
 - support for large mbox files over 4Gb;
 - find text in subject, sender. Filter by dates: a * will display all messages between the selected dates;
 - mails with attachments are tagged accordingly;
 - the preview folder is not deleted anymore;
 - program saves the scan result for faster loading;

from 1.0.0.5 to 1.0.0.6:

 - feature: added .eml file support
 - feature: added recent directoryes in file menu
 - fix: corrupt mails not shown anymore;
 - fix: correct extension was not generated for ie when content-type was not lower case;

From 1.0.0.4 to 1.0.0.5:

 - bug fix: body text was not correctly shown for some messages;

From 1.0.0.3 to 1.0.0.4:

 - lots of speed improvements on mbox parsing and list population;
 - fixed memory leak in mail view;
 - small bug fix;

From 1.0.0.2 to 1.0.0.3:

 - 2005.12.05 added toolbar buttons [eneam];
 - 2005.12.05 fixed search bug crashing program if last element was found [eneam];
 - 2005.12.01 selecting root folder clears list [eneam];
 - 2005.12.01 search bug hanged program. fixed.[eneam];

 