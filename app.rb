require 'google_drive'
require 'paypal-sdk-rest'
require 'yaml'
include PayPal::SDK::REST

class Account
  def initialize(parameters)
    paypal = parameters['Paypal']
    PayPal::SDK.configure(
      mode:          "live", # "sandbox" or "live"
      client_id:     paypal['client id'],
      client_secret: paypal['client secret key'])
  end
  
  def payments
    Payment.all.payments
  end
end

class Spreadsheet
  def initialize(parameters)
    gdrive  = parameters['Google Drive']
    session = GoogleDrive.login(gdrive['username'], gdrive['password'])
    @ws     = session.spreadsheet_by_key(gdrive['spreadsheet key']).worksheets[0]
  end
  
  def dedupe(payments)
    new_payments = payments
    
    # Drop payment if it already exists in spreadsheet...
    for row in 2..@ws.num_rows
      # ...but not if it exists with a different state
      new_payments = new_payments.delete_if { |payment| payment.id == @ws[row, 1] && payment.state == @ws[row, 4] }
    end
    
    return new_payments
  end
  
  def first_empty_row
    @ws.num_rows + 1
  end
  
  def update_with(account)
    row = first_empty_row
    
    # Add headers if they don't exist
    if row = 1
      @ws[row, 1] = "ID"
      @ws[row, 2] = "Created At"
      @ws[row, 3] = "Updated At"
      @ws[row, 4] = "State"
      @ws[row, 5] = "Payment Method"
      @ws[row, 6] = "Amount"
      @ws[row, 7] = "Description"
      @ws[row, 8] = "Email"
      @ws[row, 9] = "First Name"
      @ws[row, 10] = "Last Name"
      @ws[row, 11] = "Phone"
      @ws[row, 12] = "Shipping Address"
    end
    
    dedupe(account.payments).each do |payment|
      payment.transactions.each do |transaction|
        @ws[row, 1] = payment.id
        @ws[row, 2] = payment.create_time
        @ws[row, 3] = payment.update_time
        @ws[row, 4] = payment.state
        
        payer = payment.payer
        @ws[row, 5] = payer.payment_method
        
        @ws[row, 6] = transaction.amount
        @ws[row, 7] = transaction.description
        
        payer_info = payer.payer_info
        @ws[row, 8] = payer_info.email
        @ws[row, 9] = payer_info.first_name
        @ws[row, 10] = payer_info.last_name
        @ws[row, 11] = payer_info.phone
        
        address = payer_info.shipping_address
        @ws[row, 12] = "#{address.line1}, #{address.line2}, "\
          "#{address.city}, #{address.state} #{address.postal_code}, "\
          "#{address.country_code}"
      
      row +=1
      end
    end
    @ws.save
  end
end

# Get parameters from config file
parameters = YAML::load_file 'config.yml'

# Set up Google Drive and Paypal objects
spreadsheet = Spreadsheet.new parameters
account     = Account.new parameters

# Update spreadsheet
spreadsheet.update_with account