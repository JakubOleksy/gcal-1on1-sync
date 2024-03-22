 There are likely various ways to set this up. I did the following in my 1:1 Google Sheet:

 1. Insert -> Image -> Insert image over cells
 2. Pick whatever image you want, I just searched for Refresh Icon on Google images.
 3. When placed where you want it, select the 3 vertical dots on the image and choose Assign Script
 4. In the text box, put the function name, updateMeetingData.
 5. On the menu, choose Extensions -> Apps Scripts
 6. + a new file, name it whatever you want.
 7. Copy the function from script.js to the file created. 

 Remember, since this is accessing your calendar, only you can run the script by clicking on the image. At my place of work, I wasn't able to do this any other way where an app could run on-behalf of me, but if you can, there are likely ways you could delegate the running of the script. In practice, I run it once a week to refresh data and that was good enough for me.

 [Here](https://docs.google.com/spreadsheets/d/1Nrm1oGeMt5y6SUN0nhBW7yH-WPl5Knnh5d2GxGJBAUM/edit?usp=sharing) is a sample template for the Google Sheet that I used
