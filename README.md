drivepal
========

This application drops PayPal transactions into a Google Spreadsheet, designed to execute on a schedule.

First, set up Google Drive and PayPal in config.yml by replacing all values with proper credentials. The key for a Google Spreadsheet file can be found in its URL, as the value for its "key" query string key.

Then, run the main file, app.rb, as often as desired with:

ruby app.rb
