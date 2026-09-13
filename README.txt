SRI RADHE ENTERPRISES - BILLING SOFTWARE
==========================================

This app runs on ONE computer (a PC or laptop), and can then be opened
and used from that computer AND from a phone on the same WiFi.
(A phone alone cannot "run" the app by itself - it needs a computer
acting as the server, the same way a WiFi router needs to be on for
your phone to get internet.)

----------------------------------------------------
PART 1 - ONE-TIME SETUP (on the computer)
----------------------------------------------------
1. Install Python 3 if not already installed: https://python.org
   (On Windows, tick "Add Python to PATH" during install.)
2. Unzip this folder somewhere easy to find (e.g. Desktop).
3. Open a terminal / command prompt inside this folder and run:
       pip install -r requirements.txt

----------------------------------------------------
PART 2 - STARTING THE APP (do this every time you want to bill)
----------------------------------------------------
EASIEST WAY:
   - Windows: double-click "start_windows.bat"
   - Mac / Linux: double-click "start_mac_linux.sh"
     (or run:  ./start_mac_linux.sh  from a terminal)

This automatically starts the server AND opens the billing page
in your browser on the computer.

MANUAL WAY (if the above doesn't work):
   Open a terminal in this folder and run:
       python app.py
   Then open a browser and go to:  http://127.0.0.1:5000

Leave the black terminal/command window OPEN while billing -
closing it stops the app. You'll see a message like this:

   On THIS computer, open:            http://127.0.0.1:5000
   On your PHONE (same WiFi), open:   http://192.168.1.xx:5000

----------------------------------------------------
PART 3 - USING IT ON YOUR PHONE
----------------------------------------------------
1. Make sure your phone is connected to the SAME WiFi network
   as the computer running the app.
2. Start the app on the computer as in Part 2, and note down the
   "On your PHONE" address it prints (looks like
   http://192.168.1.XX:5000 - the numbers will vary).
3. On the phone, open Chrome/Safari and type that address in.
4. The billing page opens - it's designed with big buttons, large
   text, and one product per card so it's easy to fill in with a
   thumb on a small screen.

MAKE IT FEEL LIKE AN APP ON THE PHONE (optional, recommended):
   - Android (Chrome): open the page, tap the 3-dot menu,
     choose "Add to Home screen".
   - iPhone (Safari): open the page, tap the Share icon,
     choose "Add to Home Screen".
   This puts a "Sri Radhe Billing" icon on the phone's home screen
   that opens straight to the billing page (still needs the
   computer/server running and both devices on the same WiFi).

NOTE ON PHONE FIREWALLS: if the phone can't connect, the computer's
firewall may be blocking it the first time - just allow access when
Windows/Mac asks "Allow this app to access the network?".

----------------------------------------------------
HOW TO USE THE APP
----------------------------------------------------
- "New Invoice" tab: fill in customer details, tap "+ Add Product"
  for each item (shown as an easy-to-fill card), then tap
  "Save Invoice & Get Bill". The bill PDF opens automatically in a
  new tab; a green "View / Print Bill PDF" button also stays on the
  page in case a pop-up is blocked.
- "Past Invoices" tab: search and reprint any old bill.

----------------------------------------------------
EDITING COMPANY DETAILS
----------------------------------------------------
Open app.py, find the "COMPANY = { ... }" section near the top, and
edit the name, GSTIN, address, mobile number, and bank details there.
These appear on every printed bill.

----------------------------------------------------
KEEPING IT RUNNING ALL DAY / ALWAYS AVAILABLE
----------------------------------------------------
For daily shop use, just leave the computer on and the app started
(Part 2) during business hours. If you'd like it to run permanently
in the background, start automatically with the computer, or be
reachable from outside the WiFi (e.g. hosted online for phone-only
access with no computer needed), a developer can set that up.
