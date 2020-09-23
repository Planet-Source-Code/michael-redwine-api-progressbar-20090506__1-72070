*****************************************************************************
* API_ProgressBar, written by Michael Redwine   (michael@phuzz.net)         *
*                                            (http://www.phuzz.net)         *
* Used self-subclassing code by Paul Caton (Paul_Caton@hotmail.com)         *
*****************************************************************************

1...What
2...Why
3...What else?
4...A note about visual styles



1.  API_ProgressBar control emulates most of the properties, methods, and
	events of Microsoft's ogirinal ProgressBar control.

2.  I wrote this control because I have an obsession with not including OCX
	files in my applications... because I prefer to create single-file
	programs without dependencies.

3.  In addition, the following are added or modified:

* Marquee - Changes the style of the ProgressBar to "calculating".  This is
	what you see when the ProgressBar has the green bar constantly going
	across, over and over and over.  This property only has an effect
	when using XP or greater and using the styles manifest (see notes).
* State - Allows you to set the state of the ProgressBar to Normal (green),
	Paused (yellow), or Error (red).  This property only has an effect
	when using Vista or greater and using the styles manifest.

4.   a) This control doesn't emulate an "XP" or "Vista" progress bar.  It
	simply incorporates sytles and other such stuff in it so when your
	program is ran in said operating system (with included manifest),
	it won't look all ghetto.

     b) Modern visual styles are not any fun to program support for.  Figured
	it out, though, I think.  First off, I included a "styles_manifest"
	resource file.  Include this in your project.  Doing so will enable
	your program to use the modern visual styles.  However, in the event
	that you or a user might be using XP instead of Vista, you need to
	do the next letter...

     c) You need to add this to your main startup form, like this:

	Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

	Private Sub Form_Initialize()
	  InitCommonControls
	End Sub

	This has nothing to do with API_ProgressBar, but everything to do
	with XP sucking.  The above is a workaround that you need to use in
	order to be able to use styles.