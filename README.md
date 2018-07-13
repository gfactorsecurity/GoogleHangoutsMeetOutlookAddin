# Google Hangouts Meet Outlook AddIn

### This is an addin written using the Office JavaScript API to allow for the creation of Google Hangouts Meet events via Outlook.

Google Hangouts Meet allows for the creation of a phone number and PIN to join a conference, in addition to a URL for a normal video chat.  As of present, there is no
official Outlook plugin from Google that supports Meet.  In order to facilitate the creation of Meet events with phone numbers and PINs, I created this addin.

When a meeting is created using this addin, a new Google Calendar event is created using the Google Calendars API, with a Hangouts Meet event attached to it.  
This event is created starting at the current time and ending an hour from the current time.  We determined that despite a calendar event needing to be created
to schedule a Hangouts Meet event, once the original calendar event is deleted, the scheduled Meet session still remains valid.  In order to not junk up Google Calendar
(for users who do not use the calendar as part of G-Suite), we delete the created calendar event.

Requirements:
 * Web server with valid SSL certificate
 * G-Suite account for users of addin (does not work with free Google accounts; Hangouts Meet is only available with paid G-Suite)
 * Hangouts Meet enabled for G-Suite (this must be enabled by your G-Suite administrator)
 * Google account (free or paid) to generate API Key and Client ID for addin

Installation Instructions:
1) Modify the files below to replace the string '<your_website_here'> with the hostname of your website (for example, www.mysite.com).  Do NOT include "https://".  **NOTE**: Site MUST be HTTPS and have a valid SSL certificate installed.
   * function-file/function-file.html
   * index.html
   * google-hangouts-meet-outlook-addin-manifest.xml

2) Modify the manifest file (google-hangouts-meet-outlook-addin-manifest.xml) to include a random GUID below the line that says "Your GUID Below".  **NOTE**: You can use the following site to generate a random GUID: https://www.guidgenerator.com/online-guid-generator.aspx

3) Create a new Google Developer project and take the API Key and Client ID created and put them into function-file/function-file.js.  Replace "<your_client_id_here" with the Client ID and replace "<your_api_key_here>" with the API Key.  **NOTE**: For help creating the project and generating the required info, see this site: https://docs.simplecalendar.io/google-api-key/

4) Upload the modified documents to your website and place them in a folder at the root of the webserver called "hangoutsmeetoutlookaddin".
NOTE: You can change the path by modifying the files from Step 1 and updating the path in every location where you placed your website name.

5) Import the addin into Outlook.  You can do this by going to OWA and selecting Manage Addins from the addins menu (gear icon).  Then you can choose to add a new addin via URL.  Enter the full URL to the google-hangouts-meet-outlook-addin-manifest.xml file.

6) The addin should appear in Outlook for Windows, Outlook for Mac, and OWA 2013/2016/365.  The icon is the green Google Hangouts logo.


Usage Instructions:
1) Create a new message or calendar appointment in Outlook.

2) Click the "Settings" button from the toolbar (if using Outlook on the desktop).  Click the green Hangouts logo icon if using OWA.

3) Click "Sign In" to log into Google and authorize access to your Google Calendar.
NOTE: You MUST log in using a paid G-Suite account.  You will get an error when creating events if you do not have a paid G-Suite account, or if your administrator has not enabled Google Hangouts Meet for your G-Suite domain.

4) Once signed in, click the "Create Event" button on the sidebar, or click the "Create Meeting" button (both buttons are from this addin; only 1 will show in OWA).


Troubleshooting:
**Ensure you have a valid (trusted) SSL certificate installed on the website hosting the addin.
**Go to the URL of index.html from your web browser to verify you can access the files in question and authorize your account.
**If buttons are greyed out when creating a new message or appointment, try disabling Outlook COM Add-Ins to see if there is a conflict.  In testing, the Citrix Sharefile Outlook addin (native, not Office.JS) conflicted.  This was resolved by installing the latest version of Sharefile, or alternatively installing the Office.JS/Office Store version of the plugin.


