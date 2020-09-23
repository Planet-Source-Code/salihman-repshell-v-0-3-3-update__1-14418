**********************************************************************
	               RepShell v 0.3.1 - made in VB6
                     Copyright (c) 2000 Salih Gunaydin
                       Co-Programmer : Koen Mannaerts
                        Bug - Hunter: Wouter Tollet
  		   Email  : wippo@antwerp.crosswinds.net
**********************************************************************

CONTENTS
--------
1)Request
2)History
3)Disclaimer


HotKeys:
    Ctrl+Alt+A      : End the program
    Windows key + S : Show Start Menu on mouse position
    Windows key + F : Show Favorites menu at mouse position
    Windows key + R : Show RunDialog


1)Request
-----------------------
First of all I want to give credit to Brian, creator of BoS. Because
it was because of his BoS that I started with RepShell, thinking the coding 
could be done better. But his graphics are the best, but that slows his 
program down incredibly, because all is done with pictureboxes and those are
very memory intensive. RepShell does all its drawing directly on the form with
API which is incredibly fast.So if there are any graphics artists out there who 
would like to help I would appreciate it very much.

REPSHELL IS AN OPEN_SOURCE PROJECT.

So the source code will be publicly available in the hopes that other 
programmers will make improvements to this program, and then relay 
them back to me so I can keep releasing new versions of the program, 
with improvements that either improve the interface or the internal 
functioning of the program. For the moment I have only one co-programmer : 
Koen Mannaerts, who has shown his capabilities, as a thinking man. 
If you are interested, then mail.

Any contributors would of course get credit for their contribution.

The eventual goal is to make a compact replacement shell (RepShell) for
explorer, that has all the functionality of explorer.exe and a multitude of improvements.


2)HISTORY
---------
  v 0.3.3
     * FIXED : After changing a color and a restart, trying to change that color again crashed the program
     * FIXED : StartMenu Foldout
     * FIXED : ChooseColor Bug
     * Added QuickExplorer Icon, added option to change QuickIcons
     * Internally started adding function for multimedia purposes, such as: Playlist, tag read/write, play mp3 and other audio formats (using MCI, no MediaPlayer control)
     * Simple functions to control one audio file at a time, playlist abilities not yet activated
     * Added MediaForm, Winamp playlists are recognized and can be played
     * All Mp3 Files can be played, except those with CRC's (i discovered this when checking the properties of the files it wouldn't play in WinAmp)
     * Volume Control, pos Control
     * Decided it would be better of as a whole new project RepAmp, I will be posting this soon

  v0.3.2 
     * Added item to context menu of desktopitems to rename them
     * Changed form design and extended settings form
     * Replaced frmRun with Windows own run dialog, can someone help me retrieve it's hwnd, so I can position it
     * Totally Enabled system menu, when clicked on taskbutton + added option of always on top
     * Threw out frmSessionInfo, but will implement the info provided here someplace else
     * Threw out half of RAS Functions, instead of overlapping Windows Function, it's now complementary
     * No longer using Logs, the writelog window is no more
     * Made an ocx of the Systray and taskbutton
     * Added arrow key navigation to the desktop icons
     * Added option "Make default Shell", when unchecked "Explorer.exe" becomes the default shell again
     * BUG : ChooseColor Dialog crashes outside the VB IDE
     * Option "Alwas on Top" in systemmenu now fully working

  v0.3.1
     * Show system menu of prog when right clicked on taskbox, but not yet executing commands
     * Replaced frmMenu by API created menus, any help with replacing the Start Menu by Api, would be very welcome
     * Working Desktop refreshing
     * Added back context-menu handlers which now work perfectly (including context menu for recycle bin)
     * Speed up file searching, Now using FindFile API
     * Added SysTray
     * FIXED : When an item was to be deleted and could not be found in the Systray, the last item got deleted.
     * FIXED : The StartMenu and the other frmStart based menus didn't get unloaded when you clicked on something else.
     * FIXED : TaskButtons got unloaded and reloaded when title changed, this caused a flicker. Now only the caption changes.
     * Internal changes to RAS Connections
     * Added Option screen
     * Added arrow key navigation to start menu
     * FIXED : Tray Icons did not get updated

  v0.3
     * Extended Email configuration. Designed new interface, internal changes to the Winsock module
     * saving emails, folders system
     * Seperated RepMail from RepShell, increase in speed and overall stability
     * Speed up Start Menu, very less memory intensive. Doesn't use any additional controls, previous versions loaded a PictureBox for every item, now everthing is drawn on form using API's --> Very Fast
     * Directory's under 'My Computer' are sorted
     * Added program icons on tasklist
     * FIXED : Taskbuttons kept switching if windows had the same name
     * Removed context-menu handlers. They weren't working perfectly and needlessly complicated the program. Will add them later. If anyone wants to help out, I would be very greatful.
     * Porting to English complete.

  v0.2
     * Some serious bugfixes.
     * Desktop refreshing possibilities, with problems: See modFolderspy.bas
     * Run - runs everything
     * Added taskbox
     * Dial Up Networking : recognition of all Dial Up, ability to add new dial up, get properties, get session info
     * Recognition of all present drives, click on 'My computer'
     * Added Mainlog : saves the messages that appear on the Main Log Window
     * Added ModemLog : saves internet connection times, etc
     * Released on PSC

  v0.1
     * Just released it to some friends, to see if it's any good.


3)DISCLAIMER
------------
THIS SOFTWARE AND THE ACCOMPANYING FILES ARE DISTRIBUTED "AS IS" AND WITHOUT WARRANTIES
AS TO PERFORMANCE OF MERCHANTABILITY OR ANY OTHER WARRANTIES WHETHER EXPRESSED OR IMPLIED.
NO WARRANTY OF FITNESS FOR A PARTICULAR PURPOSE IS OFFERED. THE USER MUST ASSUME THE ENTIRE
RISK OF USING THE SOFTWARE.