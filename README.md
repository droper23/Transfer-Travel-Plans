The purpose of this code is to use Google Maps to quickly and automatically find and save travel plans to and from cities. 
Originally this was used for LDS missionary transfers in Hungary, where missionaries would come into Budapest from their seperate cities and then leave from Budapest to their new city.
This code was created to automate the finding of travel plans for these missionaries, along with taking screenshots of the travel plans in order to send the plans directly to the missionaries.

There are really three main parts to the code:
  1. Using a custom UI to change the date and time of travel, as well as who is moving from where to where
  2. Pulling up travel plans automatically using Google Maps on a browser
  3. Saving the data in an Excel spreadsheet and the screenshots in a seperate folder

In order for the code to work, simply run the code. A window will pop up with the options to change time and day of transfers, as well as adding personnel to be transferred and the to and from cities.
As of right now, the list of possible people and transfers have to be manually written into the code, but it's pretty simple to change the variables to pull from an Excel sheet or something similar.
After clicking the "save transfers" button, sit back and relax. Browser windows will pop up, maps will load, and after it's done, the data will be saved to a folder labelled "GoogleMapsTrains". 
