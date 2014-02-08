require 'google_drive'
require 'paypal-sdk-rest'
require 'yaml'
include PayPal::SDK::REST

# Get parameters from config file
parameters = YAML::load_file 'config/parameters.yml'

# Set up Google Drive
gdrive  = parameters['Google Drive']
session = GoogleDrive.login(gdrive['username'], gdrive['password'])
ws      = session.spreadsheet_by_key(gdrive['spreadsheet key']).worksheets[0]

# Set up Paypal
paypal = parameters['Paypal']
PayPal::SDK.configure(
    mode:          "sandbox", # "sandbox" or "live"
    client_id:     paypal['client id'],
    client_secret: paypal['client secret key'])

# Push data
row = ws.num_rows
Payment.all.payments.each do |payment|
  row += 1
  ws[row, 1] = payment.id
end
ws.save