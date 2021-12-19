
Changes
---

v 1.0.3.25

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for Ctrl-C to Copy text in the message pane.;
 - Added code to check whether a user has permission to write to registry. Mbox Viewer will not work if the user can’t write to Window’s registry.;
 - Changed the file name of exported mail as eml file from message.eml to mime-message.eml to minimize potential conflict with mail’s attachment name.;
 - Updated User Manual to describe new and updated features.;

v 1.0.3.24

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for Gmail Labels. User must run a separate step on active mail archive to recreate Gmail labels. See User Guide for details and limitations to be addressed in the next realese.;
 - Disable "Create Folder" unfinished option left enabled by mistake while working on Gmail Labels support;
 - Updated User Manual to describe new and updated features;

v 1.0.3.23

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for Gmail Labels. User must run a separate step on active mail archive to recreate Gmail labels. See User Guide for details and limitations to be addressed in the next realese.;
 - Updated User Manual to describe new and updated features;

v 1.0.3.22

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for Gmail Labels. User must run a separate step on active mail archive to recreate Gmail labels. See User Guide for details and limitations to be addressed in the next realese.;
 - Updated User Manual to describe new and updated features;

v 1.0.3.21

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for forwarding of mails directly form MBox Viewer via Gmail, Outlook or Yahoo SMTP server;
 - Updated User Manual to describe new and updated features;

v 1.0.3.20

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for read only media such as CDs;
 - Updated User Manual to describe new and updated features;

v 1.0.3.19

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added option to print mails directly to PDF with no page header and footer via Microsoft Edge browser;
 - Fixed merging files via file list and added support for wildcard file names;
 - Enhanced the stack trace dump in Exception;
 - Updated User Manual to describe new and updated features;

v 1.0.3.18

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Enhanced Export All Mail Attachments to include inline and optionally embeded attachments;
 - Updated User Manual to describe new and updated features;

v 1.0.3.17

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added -EML_PREVIEW_MODE command line option to hide Mbox Tree and Mail List panes when -MAIL_FILE is configured. ESCAPE key will terminate the mbox viewer;
 - Added -MBOX_MERGE_LIST_FILE and -MBOX_MERGE_TO_FILE command line options to merge mbox files listed in the file. The merged file is automatically open in mbox viewer;
 - Fixed rare crash when Removing Duplicate mails with extremly large number of duplicate mails;
 - Updated User Manual to describe new and updated features;

v 1.0.3.16

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - It is RECOMMENDED to upgrade to this release if you expierence unexpected regeneration of index files;
 - Added option to hide mbox files from the tree view and restore these files when needed;
 - Enhanced "Open File" option after creation of archive file is done to avoid exceeding the file name length limit (currently 260);
 - Fixed missing blank character when unfolding field text lines;
 - Updated index file version due to unfolding issue;
 - Updated User Manual to describe new and updated features;

v 1.0.3.15

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added additional filter rules to redesigned Find Advanced dialog;
 - Added option to show expanded list of header fields in the message header pane;
 - Added option to show raw mail header fields;
 - Added option to restore message hints;
 - Updated User Manual to describe new and updated features;

v 1.0.3.14

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Enhanced to create missing index files for mbox files larger than 100Gb;
 - Enhanced to make windows placements persists across multiple restart;
 - Added option to export all mail attachments;
 - Added option to export all mails as separate Eml mails;
 - Added option to configure attachment name separator, including CRLF, in CSV file;
 - Added option to list all attachment names in TEXT, HTML and PDF files. In HTML files, Names are links to attachment files;
 - Improved detection of file extensions of attachment files;
 - Enhanced mail grouping by conversation for Gmail file types with X-GM-THRID field present;
 - Enhanced mail grouping by conversation for non Gmail file types when sent mails are missing;
 - Made default Select Folder dialog resizable;
 - Updated User Manual to describe new and updated features;

v 1.0.3.13

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added button to toolbar to hide/unhide the Mail Tree view;
 - Added check to validate path of the folder housing mbox files;
 - Updated User Manual to describe new and updated features;

v 1.0.3.12

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added support for multiple folders housing mbox files under the Mail Tree view;
 - Fixed HTML/PDF mail header dialog. User was able to configure custom font while the default font is set;
 - Updated User Manual to describe new and updated features;

v 1.0.3.11

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added capability to customize the HTML/PDF mail header output;
 - Fixed incorrect header field text in HTML/PDF showing up occasionally;
 - Added HTML2PDF-single-chrome-canary.cmd, HTML2PDF-all-chrome-canary.cmd and HTML2PDF-group-chrome-canary.cmd scripts to remove header and footer in PDF;
 - Updated User Manual to describe new and updated features;

v 1.0.3.10

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Relaxed the mbox parsing rules in order to accept mbox files with different From_ line format that indicates starts of new mail;
 - Corrected mail content when exporting to eml file;
 - Updated User Manual to describe new and updated features;

v 1.0.3.9

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Fixed the rare crash when parsing mbox file. Unverified fix submitted in the v1.0.3.8 was not fully reliable. Upgrade to v1.0.3.9 is recommended;
 - Added capability to customize the background color of most of the display panes;
 - Updated User Manual to describe new and updated features;

v 1.0.3.8

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Enhanced attachment window to support Unicode UTF16 names of non-inline attachments;
 - Implemented the number of small code fixes;
 - Bumped CACHE_VERSION to 17;
 - Updated User Manual to describe new and updated features;

v 1.0.3.7

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Submitted unverfied fix to hopefully resolve rare crash when parsing mbox file;
 - Added logging to a file of two raw mails upon catching exception;
 - Added logging of the stack trace upon catching exception by the separate instrumented version of the mbox viewer;
 - Added HELP.txt to describe enhancements to help to speed up the resolution of execution exceptions;
 - Enhanced performance of internal Picture Viewer when resizing images;
 - Replaced the star * character as the attachment indicator with the industry standard paper clip;
 - Enhanced to assure that the Hint Message Box always stays on top of the main window;
 - Enhanced Print Config Dialog to allow users to customize file name template;
 - Updated User Manual to describe new and updated features;

v 1.0.3.6

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Enhanced the internal Picture Viewer to observe picture size as set in the meta data in the image file;
 - Added Attachment Config dialog to configure attachment window size, content and attachment indicator criteria;
 - Added startup check to warn user that the instance of mbox viewer might be running already;
 - Fixed crash when parsing mbox with mails older than epoch time;
 - Resolved some issues with Date Filter in search mode;
 - Updated mbox file parsing and bumped CACHE_VERSION to 16;
 - Updated User Manual to describe new and updated features;

v 1.0.3.5

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Fix to limit the text length when printing of Message Text column in the CSV file for emails with Html TEXT block only and no plain TEXT block;
 - Added context level hints on how to use mbox viewer;
 - Added Attachment Name field to Find and Advanced Find filters;
 - Added option to Find and Advanced Find dialogs to find mails that don't much search criteria;
 - Enhanced handling of '*' search string in Find and Advanced Find dialogs to be able to find subset of mails that have CC, BCC and mails with attachments;
 - Updated mbox file parsing and bumped CACHE_VERSION to 15;
 - Added option to Print Config dialog to control mail body background color with wkhtmltopdf tool;
 - Updated User Manual to describe new and updated features;

v 1.0.3.4

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added LICENSE.txt file;
 - Enhanced handling of MAIL_FILE command line option;
 - Fixed printing of Message Text column in the CSV file for emails with Html TEXT block only and no plain TEXT block;
 - Fixed crash due to very large to, cc, bcc header fields greater than 10000 characters when printing to CSV file;
 - Added new column to CSV file to export Attachment Names;
 - Resolved Date format in the filename when printing to PDF file;
 - Updated User Manual to describe new and updated features;

v 1.0.3.3

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added an option to create inline image cache on the fly to reduce initial startup time;
 - Enhanced mail header field decoder according to RFC 2047 to resolve occasional text display issues;
 - Updated User Manual to describe new and updated features;

v 1.0.3.2

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Enhanced handling of missing character set in the raw email that could result in improper text display;
 - Added Mail Retrival In Progress Status text in the Status Bar;
 - Added Mail Archive Open in Progress Status text in the Status Bar;
 - Added Sort Column in Progress Status text in the Status Bar;
 - Enhanced handling of inline images;
 - Updated User Manual to describe new and updated features;

v 1.0.3.1

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - RESOLVED Out of Memory crash due to corrupted mboxview index file created during initial parsing of mail archive;
 - Added capability to merge/concatenate multiple archive files and remove duplicate mails;
 - Added option to print directly to PDF without user interaction by leveraging external HTML to PDF converter tools such as Chrome Browser or wkhtmltopdf;
 - Added option to print each mail to a separate PDF file;
 - Added scripts to merge per mail PDF files to support printing large number of mails to a single PDF file;
 - Added vertical bar in the first column to mark mails that are also on the User Selected Mails list;
 - Added handling of CTRL-A to select all mails in the Summary Window. Windows must be selected first;
 - Added option to insert Page Break after each mail when printing multiple mails to single Text or PDF file;
 - Added option to remove/add the background color to the mail header;
 - Added options to export CC and BCC header fields to CSV spreadsheet file;
 - Added CC and BCC fields to a search criteria in Find Advanced and Find dialog;
 - Added option to limit the maximum text length when printing mail text to CSV;
 - Updated Help options;
 - Updated User Manual to describe new and updated features;

 v 1.0.3.0

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added mail list editing capability to enable users to compose list as subset of mails from multiple searches;
 - Added capability to archive search results and/or mail list composed by users to file;
 - Added capability to reload mail list composed by users from archive;
 - Added capability to select one or more mails and print, remove and/or copy;
 - Added option to directly print mails to user selected printer;
 - Added Advanced Find/Search capability;
 - Added option to find all matching mails;
 - Added option to set Message Window position to Bottom, Right or Left;
 - Added option to configure file name for printing a mail to a file;
 - Updated User Manual to describe new and updated features;

v 1.0.2.8

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Fixed code to address random startup failures;
 - Added handling of embedded images declared incorrectly as non inline attachments in the mail;
 - Fixed incorrect text alignment in some cases when printing group of mails to HTML;
 - Added html2text conversion to print mails without plain text block;
 - Added progress bar when printing mail archive to Text or HTML or CSV files;
 - Fixed race condition (introduced in 1.0.2.7) that could result in message and mail summary out of sync;
 - Updated User Manual to describe new and updated features;

v 1.0.2.7

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Fixed picture orientation in Picture Viewer;
 - Improved performance of Picture Viewer to reduce flicker;
 - Added zoom slider and image dragging using mouse to Picture Viewer;
 - Added mail message retrieval state in the status bar;
 - Updated User Manual to describe new and updated features;

v 1.0.2.6

 - To improve mboxview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - Added options to print all mails to CSV or Text or HTML files;
 - Added options print single mail to Text or HTML file;
 - Added option to group all related mails in the summary window;
 - Added option to print related mails/mail group to Text or Html file;
 - Added handling of inline/embeded images;
 - Added option to open single mail or group of related mails in external Browser;
 - Added option to find user defined text and highlight all occurences;
 - Updated User Manual to describe new and updated features;
 - Synchronized User Manual version number with the mboxview release number

v 1.0.2.5

 - To improve mbxoview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - submitted Picture Viewer source code files to Github, missing in v1.0.2.4;
 - users not interesetd in the source code don't need to reinstall mboxview binary file;

v 1.0.2.4

 - To improve mbxoview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - added picture viewer dialog to preview attachments with the png,jpg,jpe,jpeg,gif,bmp,ico,tif,tiff,jfif,emf,wmf,dib suffixes;
 - all attachements are still accessible by double clicking selected email in the summary window or by selecting View->"View EML"menu option;
 - corrected small memory leak;
 - updated User Manual to describe new and updated features;

v 1.0.2.3

 - To improve mbxoview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - In this release, v1.0.2.3, incremented the software and mboxview index file versions only;
 - Unfortunatelly, due to mistake, the software and mboxview index file versions were not incremented in v1.0.2.2;
 - The mboxview file generated by v1.0.2.1 will be processed correctly by v1.0.2.2, however message search performance will not be optimal;
 - The mboxview file generated by v1.0.2.2 is not backward compatible and will crash v1.0.2.1;
 - Preferred solution is to install v1.0.2.3 or manaully delete existing mboxview file(s) (not recommended);

v 1.0.2.2

 - To improve mbxoview, please post reviews on what works, what doesn't, create bug tickets and enhancment requests;
 - updated User Manual to describe new and updated features;
 - added Next/Previous search directions options to the Find dialog;
 - reduced number of false positive searches;
 - fixed potential search issue introduced in v 1.0.2.1;

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
