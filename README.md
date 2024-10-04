# raid-statistics
This script ist made for Google Spreadsheets, to help you and your static to analyse your weekly raid and strike runs visualized in days.

following need to be done
## How to Install
### Get the Script into a google Spreadsheet
1. Download the latest raid-statistics.zip here: [Releases](https://github.com/Darkister/raid-statistics/releases)
2. Unzip the Zip-File
3. Open the unziped folder. There is the File 'main.js' in it.
4. Open the File 'main.js' in any editor of your choice. Simple Notepad is absolutly fine.
5. Visit [Google Spreadsheets](https://docs.google.com/spreadsheets/) and if not logged in already, login to your google account.
6. Create a new empty spreadsheet.
7. Inside your new spreadsheet click on "Erweiterungen" ("Extensions"?) -> Apps Script, a new tab should open
8. delete what ever is inside the default file
9. Copy&Paste the Code from the previously downloaded 'main.js' File into your Script Editor
10. Save the file inside your Script Editor 

### First steps and Permissions
1. Inside the Script Editor you can now run Functions of the Script
2. Make sure that the function 'createFullLayout' is selected in the dropdown, press "Ausführen" ("Run"?)
3. You need to give Permissions to your script, just follow the instructions on the Screen
4. At one Point there is a Red Triangle and bit lower a small gray text "Erweitert" ("Advanced"), click on it and click on "Open Project (unsafe)" -> Aggree on the next Screen
5. Wait until the process is done, it could need up to 5 minutes

Now have a look at your Spreadsheet, the basic layout should be created now.

### How to use as a user
You should be done, Happy logging.

I'm just kidding, there will be a couple of How-Tos comming soon in on the wiki page which will be linked here -> placeholder

If you don't want to read thousand of pages a long story short
- go to the tab "Settings"
- adjust the "players to view" if needed
- go to the tab "Setup und Co"
- enter the Accountnames of the players to view -> should be you static members (alt accounts currently not supported)
- go back to the tab "Settings"
- double-click the gray box next to "Enter Logs here:"
- paste your logs here -> be careful with the amount of logs, i recommend to only paste in 10 logs at a time
- wait until calculation is done

## How to use for developer
Clone the Repository with an IDE of your choice, personally I prefer VS-Code, but others should also work

Run
```
npm install
npm i @google/clasp -g
```
to install all needed packages.

Make familiar with clasp [Working with Google Apps Script in Visual Studio Code using clasp](https://yagisanatode.com/2019/04/01/working-with-google-apps-script-in-visual-studio-code-using-clasp/)

## Further documentations
* [Permissions for trigger functions](https://stackoverflow.com/questions/58359417/you-do-not-have-permission-to-call-urlfetchapp-fetch)

## Contact
* Mail - darkisters.world@gmail.com
* Discord - darkister
* Visit my own DC-Server - [Darkisters World Community Server](https://discord.gg/wMuQnYVNTv) -> mainly in german, but give your self the Role "Coding Stuff" in the channel "verwalte-deine-rolle"
* Guild Wars 2 - blackicedragon.3579
* Twitch - [Darkister](https://www.twitch.tv/darkister)

## Special Thanks
To my Raid Static LGWee
* Nekrom Sykox.5720 -> Thanks to the Leader of the static and for all the awesome functionality ideas