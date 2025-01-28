require 'groq'
require 'json'
require 'roo'  # Add Roo gem for reading Excel files

# Set default model ID
model_id = 'llama-3.3-70b-versatile'

# Ask user for the Groq API key
puts "Please enter your Groq API key:"
api_key = gets.chomp.strip

# Validate API key (basic check)
if api_key.empty?
  puts "API key cannot be empty."
  exit
end

# Initialize the Groq client
client = Groq::Client.new(api_key: api_key, model_id: model_id)

# Function to read existing classifications from a file
def read_existing_classifications(filename = 'classifications.txt')
  classifications = {}
  if File.exist?(filename)
    File.readlines(filename).each do |line|
      product_name, classification = line.split(':').map(&:strip)
      classifications[product_name] = classification
    end
  end
  classifications
end

# Function to classify multiple products with dynamic classifications
def classify_products(client, product_names)
  # Ensure all product names are in UTF-8 encoding
  product_names.map! { |name| name.encode('UTF-8', invalid: :replace, undef: :replace, replace: '') }

  messages = [{
    role: 'user',
    content: "Please classify these products. The products are: #{product_names.join(', ')}. Respond only with a valid JSON object like this: {\"classifications\": {\"product_name\": \"classification\"}}."
  }]
  
  begin
    response = client.chat(messages)

    if response && response["content"]
      raw_content = response["content"].strip
      puts "Raw Response: #{raw_content.inspect}"

      ai_response = JSON.parse(raw_content)

      if ai_response.is_a?(Hash) && ai_response.key?("classifications")
        classifications = ai_response["classifications"]
        return classifications
      else
        return "Error: Response format is incorrect."
      end
    else
      return "No content received from AI."
    end
    
  rescue JSON::ParserError => e
    puts "Error parsing response: #{e.message}"
    puts "Raw response was not valid JSON: #{raw_content}"
    return "Error in classification."
    
  rescue StandardError => e
    puts "An error occurred while communicating with the AI: #{e.message}"
    return "Error in classification."
  end
end

# Function to save product name and classification to a file
def save_to_file(product_name, classification, filename = 'classifications.txt')
  File.open(filename, 'a') do |file|
    file.puts("#{product_name}: #{classification}")
  end
  puts "Saved '#{product_name}' classified as '#{classification}' to #{filename}."
end

# Function to delete all classifications from the file
def delete_classifications(filename = 'classifications.txt')
  File.open(filename, 'w') {} # Truncate the file to delete all contents
  puts "All classifications have been deleted from #{filename}."
end

# Function to classify products from an Excel file
def classify_products_from_excel(client)
  puts "Please enter the name of your Excel file (with extension, e.g., 'products.xlsx'):"
  file_path = gets.chomp.strip

  # Open the Excel file using Roo gem
  begin
    xlsx = Roo::Spreadsheet.open(file_path)
  rescue => e
    puts "Error opening file: #{e.message}"
    return
  end

  # Automatically use the first sheet
  xlsx.default_sheet = xlsx.sheets.first
  puts "Using sheet: #{xlsx.default_sheet}"

  # List available columns and ask for the column number
  columns = ('A'..'Z').to_a[0...xlsx.last_column] # Adjust based on actual number of columns in use
  puts "Available columns: #{columns.join(', ')}"
  
  puts "Please enter the number corresponding to the column you want to use (e.g., A = 1):"
  column_number = gets.chomp.to_i

  # Validate column number
  if column_number < 1 || column_number > columns.size
    puts "Invalid column number. Please enter a number between 1 and #{columns.size}."
    return
  end

  # Get the corresponding column letter
  column_letter = columns[column_number - 1]

  # Extract unique product names from the designated column starting from row 2 (ignoring header)
  product_names = xlsx.column(column_letter)[1..-1].compact.uniq
  
  # Read existing classifications to filter out already classified products
  existing_classifications = read_existing_classifications()
  
  # Filter out already classified products and keep only unique ones without classifications
  new_product_names = product_names.reject { |name| existing_classifications.key?(name) }

  if new_product_names.empty?
    puts "All products are already classified. No new classifications needed."
    return
  end
  
  # Classify all unique product names using AI and save results to file with dynamic classifications.
  classifications = classify_products(client, new_product_names)

  if classifications.is_a?(Hash)
    classifications.each do |product_name, classification|
      save_to_file(product_name, classification)
    end
    
    puts "All unique product names have been classified."
  else
    puts classifications # Print error message if classification failed.
  end
end

# Function to list all classified products and their respective classifications.
def list_all_classified_products()
   existing_classifications = read_existing_classifications()
   if existing_classifications.empty?
     puts "No classified products found."
   else 
     puts "Classified Products:"
     existing_classifications.each do |product_name, classification|
       puts "#{product_name}: #{classification}"
     end 
   end 
end 

# Function to list all products and their classifications from an Excel file 
def list_products_and_classifications_from_excel()
   puts "Please enter the name of your Excel file (with extension, e.g., 'products.xlsx'):"
   file_path = gets.chomp.strip

   begin 
     xlsx = Roo::Spreadsheet.open(file_path)
   rescue => e 
     puts "Error opening file: #{e.message}"
     return 
   end 

   xlsx.default_sheet = xlsx.sheets.first 
   puts "Using sheet: #{xlsx.default_sheet}"

   columns = ('A'..'Z').to_a[0...xlsx.last_column]
   puts "Available columns: #{columns.join(', ')}"

   puts "Please enter the number corresponding to the column you want to use (e.g., A = 1):"
   column_number = gets.chomp.to_i

   if column_number < 1 || column_number > columns.size 
     puts "Invalid column number. Please enter a number between 1 and #{columns.size}."
     return 
   end 

   column_letter = columns[column_number - 1]
   
   # Extract product names starting from row two.
   product_names = xlsx.column(column_letter)[1..-1].compact.uniq
   
   existing_classifications = read_existing_classifications()

   # Display each product with its classification.
   product_names.each do |product_name|
     classification = existing_classifications[product_name] || "Not classified"
     puts "#{product_name}: #{classification}"
   end 
end 

# Main menu loop 
loop do 
   puts "\nMenu:"
   puts "0. Classify a single product"
   puts "1. Classify products from an Excel file"
   puts "2. List all products and their classifications from an Excel file"
   puts "3. List all classified products and their respective classifications"
   puts "4. Delete all classifications"
   puts "5. Exit"

   print "Choose an option (0/1/2/3/4/5): "
   option = gets.chomp.strip

   case option 
     when '0'
       print "Enter a product name (or type 'exit' to quit): "
       product_name = gets.chomp.strip
        
       break if product_name.downcase == 'exit'
        
       existing_classifications = read_existing_classifications()
        
       if existing_classifications.key?(product_name)
         puts "#{product_name} is already classified as '#{existing_classifications[product_name]}'."
         next 
       end
        
       classification = classify_products(client, [product_name]) # Use dynamic classification here
        
       if classification.is_a?(Hash) && classification.key?(product_name)
         save_to_file(product_name, classification[product_name])
       else 
         puts "Error classifying product."
       end
        
     when '1'
       classify_products_from_excel(client)

     when '2'               # Option for listing products and their classifications from an Excel file.
       list_products_and_classifications_from_excel()

     when '3'               # New case for listing all classified products.
       list_all_classified_products()

     when '4'               # Updated case for deleting classifications.
       delete_classifications()

     when '5'               # Updated exit option.
       puts "Exiting..."
       break

     else 
       puts "Invalid option selected. Please try again." 
   end 
end 

puts "Program ended."
