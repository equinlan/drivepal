require 'google_drive'
require 'paypal-sdk-rest'
require 'yaml'
include PayPal::SDK::REST

class Account
  def initialize(parameters)
    paypal = parameters['Paypal']
    PayPal::SDK.configure(
      mode:          "sandbox", # "sandbox" or "live"
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
    if row == 1
      headers = ["ID", "Created At", "State", "Payment Method",
        "Amount", "Currency", "Description", "Email", "First Name",
        "Last Name", "Phone"]
      for col in 1..headers.size
        @ws[row, col] = headers[col - 1]
      end
      
      row += 1
    end
    
    dedupe(account.payments).each do |payment|
      payment.transactions.each do |transaction|
        
        # Some typing shortcuts
        payer = payment.payer
        payer_info = payer.payer_info
        amount = transaction.amount
        
        # Prepare the data to appear in columns
        data = [payment.id, payment.create_time, payment.state,
          payer.payment_method, amount.total, amount.currency,
          transaction.description, payer_info.email,
          payer_info.first_name, payer_info.last_name, payer_info.phone]
        
        # Print the data for the current row
        for col in 1..data.size
          @ws[row, col] = data[col - 1] || "N/A"
        end
      
        # Next row
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