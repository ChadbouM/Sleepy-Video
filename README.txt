Sleepy-Video

A utility for playing videos with a shutdown timer.
Includes installation features when running the script 
without tags.

Author:      Michael Chadbourne
Last Edited: 4/13/16
Version:     1.0.0

PREREQUISITES:
 * PowerShell 3+

INSTALLATION:
 * Run the script, accepting the prompt for Administrator permissions
 * If the script is not installed, it will prompt you to install.
 * Once the script is installed, running it will bring up the 
     activation prompt.
 * From the activation prompt. Hit '1' To Activate the
     Context menu item: "Watch Now Then Shutdown"
	 It will prompt for the menu-text before activating.
 * From the activation prompt. Hit '2' To Activate the
     Context menu item: "Watch at Bedtime"
	 It will prompt for the menu-text before activating.
 * From the activation prompt. Hit 'u' To Uninstall the 
     Script, deleting the install-copy, and removing the 
	 registry keys associated with the install and the 
	 active menu items.
	 
UPDATE:
 * To update the script, run an updated version of the script, and follow the
 Prompt, to Update to that script-version.

USE:
 * Active Context menu items, will appear when right-clicking a file.
 * Right click a .MP4 or a .MKV file, and select the appropriate menu entry.
 * Watch Now will play the video immediatley, shutting down a minute
    after the video has ended.
 * Bedtime Mode will prompt for when the computer should shut down.
     playing the selected video so that it will end one minute before
	 the timed shutdown. If there is not enough time to play the video
	 before shutdown, it will start the video immediatly, shutting down
	 at the given time (Mid-Video).
	 
COMMAND-LINE:
 * To use this script on the command line:
     Run the script with the following signature:
	 > Sleepy-Video.Ps1 <Directory> ([-Run] || [-Timed])
	 
	 Where <Directory> is the path to the video you wish to play.
	 The -Run Tag will play the video in Watch Now mode.
	 The -Timed Tag will prompt for a shutdown time.
	 
	 Running the script without any Tags will result in the 
	 install/activation prompts.
	 