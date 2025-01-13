require 'roo'
require 'write_xlsx'
require 'levenshtein'

# Predefined product list
PRODUCT_LIST = []

def get_input_file(default_directory)
  loop do
    print "Enter the name of the input Excel file (with .xlsx extension): "
    input_file = gets.chomp
    full_path = File.join(default_directory, input_file)
    begin
      Roo::Spreadsheet.open(full_path)
      return full_path
    rescue Errno::ENOENT
      puts "File '#{full_path}' not found. Please try again."
    end
  end
end

def get_columns_to_clean
  loop do
    begin
      print "Enter the column indices to clean (comma-separated, e.g., 1,2 for columns A and B): "
      columns_input = gets.chomp
      columns = columns_input.split(',').map(&:strip).map(&:to_i)
      if columns.any? { |col| col < 1 }
        puts "Column index must be at least 1."
        next
      end
      return columns
    rescue ArgumentError
      puts "Invalid input. Please enter valid integers."
    end
  end
end

def get_output_file_details(default_directory)
  print "Enter the output directory (or '0' to use '#{default_directory}'): "
  output_directory = gets.chomp
  
  output_directory = default_directory if output_directory == '0'
  
  print "Enter the output file name (without extension): "
  output_file_name = gets.chomp
  
  File.join(output_directory, "#{output_file_name}.xlsx")
end

def similarity_ratio(a, b)
  max_length = [a.length, b.length].max
  return 0.0 if max_length == 0
  distance = Levenshtein.distance(a, b).to_f
  similarity = 1 - (distance / max_length)
  similarity
end

def get_sorted_matches(input_text, products, n=5)
  similarities = products.map { |p| [p, similarity_ratio(input_text, p)] }
  similarities.sort_by { |x| -x[1] }.take(n)
end

def ask_for_confirmation(input_text, match, similarity)
  puts "\nOriginal product name: #{input_text}"
  puts "Suggested change to: #{match}"
  puts "Similarity: #{(similarity * 100).round(1)}%"
  
  loop do
    print "Accept this change? (Y/N): "
    response = gets.chomp.upcase
    return response == 'Y' if ['Y', 'N'].include?(response)
    puts "Please enter Y or N"
  end
end

def select_from_list(input_text, matches)
  puts "\nOriginal product name: '#{input_text}'"
  puts "Please select the correct product name from the following options:"
  puts "0. Keep original value (no change)"
  matches.each_with_index do |(product, similarity), i|
    puts "#{i + 1}. #{product} (Similarity: #{(similarity * 100).round(1)}%)"
  end
  
  loop do
    begin
      print "\nPlease select the correct product number (0-#{matches.length}): "
      choice = gets.chomp.to_i
      if (0..matches.length).include?(choice)
        if choice == 0
          puts "Keeping original value: #{input_text}"
          return nil
        else
          puts "Selected: #{matches[choice-1][0]}"
          return matches[choice-1][0]
        end
      end
    rescue ArgumentError
    end
    puts "Please enter a number between 0 and #{matches.length}"
  end
end

def find_closest_match(input_text, product_list, threshold=0.6)
  return nil if input_text.nil? || !input_text.is_a?(String)
  
  cleaned_text = input_text.strip
  return nil if cleaned_text.empty?
  
  matches = get_sorted_matches(cleaned_text, product_list)
  return nil if matches.empty?
  
  best_match, similarity = matches[0]
  
  if similarity >= 0.85
    puts "\nMatch found: #{cleaned_text}"
    return best_match
  end
  
  if similarity >= threshold
    return best_match if ask_for_confirmation(cleaned_text, best_match, similarity)
  end
  
  select_from_list(cleaned_text, matches)
end

def get_product_list_file(default_directory)
  loop do
    print "Enter the name of the product list file (with .txt extension): "
    input_file = gets.chomp
    full_path = File.join(default_directory, input_file)
    begin
      products = File.readlines(full_path, encoding: 'UTF-8').map(&:strip).map { |line| line.gsub(/^"|"$/, '') }.reject(&:empty?)
      puts "Warning: File is empty" if products.empty?
      return [full_path, products]
    rescue Errno::ENOENT
      puts "File '#{full_path}' not found. Please try again."
    rescue StandardError => e
      puts "Error reading file: #{e.message}"
    end
  end
end

def load_product_list(filename)
  begin
    products = File.readlines(filename, encoding: 'UTF-8').map(&:strip).map { |line| line.gsub(/^"|"$/, '') }.reject(&:empty?)
    return products
  rescue Errno::ENOENT
    puts "Warning: #{filename} not found. Creating new file."
    File.write(filename, '')
    return []
  end
end

def save_product_list(new_products, filename)
  begin
    existing_products = File.readlines(filename, encoding: 'UTF-8').map(&:strip).map { |line| line.gsub(/^"|"$/, '') }.reject(&:empty?)
  rescue Errno::ENOENT
    existing_products = []
  end

  all_products = existing_products.dup
  new_products.each do |product|
    all_products << product unless existing_products.include?(product)
  end

  File.open(filename, 'w:UTF-8') do |f|
    all_products.each do |product|
      f.puts "\"#{product}\""
    end
  end
end

def handle_unmatched_items(unmatched_items, product_list_file)
  return if unmatched_items.empty?
  
  puts "\nProcessing unmatched items:"
  new_products = Set.new
  
  unmatched_items.sort.each do |item|
    loop do
      print "\nAdd \"#{item}\" to product list? (Y/N): "
      response = gets.chomp.upcase
      if ['Y', 'N'].include?(response)
        if response == 'Y'
          new_products.add(item)
          puts "Added \"#{item}\" to product list"
        end
        break
      end
      puts "Please enter Y or N"
    end
  end
  
  if !new_products.empty?
    save_product_list(new_products, product_list_file)
    puts "\nUpdated product list saved to #{product_list_file}"
  end
end

def clean_product_names(input_file, column_indices, output_file, product_list_file, product_list)
  workbook = Roo::Spreadsheet.open(input_file)
  sheet = workbook.sheet(0)
  
  changes_made = 0
  total_processed = 0
  unmatched_items = Set.new
  confirmed_matches = {}

  # Create new workbook for output
  output_workbook = WriteXLSX.new(output_file)
  output_worksheet = output_workbook.add_worksheet

  # Copy all data to new workbook
  (1..sheet.last_row).each do |row|
    (1..sheet.last_column).each do |col|
      output_worksheet.write(row-1, col-1, sheet.cell(row, col))
    end
  end

  column_indices.each do |col|
    # Start from row 2 (skip header)
    (2..sheet.last_row).each do |row|
      cell_value = sheet.cell(row, col)
      if cell_value.is_a?(String)
        total_processed += 1
        original_value = cell_value.strip
        puts "\nProcessing row #{row}, column #{col}: #{original_value}"
        
        if confirmed_matches.key?(original_value)
          matched_product = confirmed_matches[original_value]
          if matched_product && cell_value != matched_product
            output_worksheet.write(row-1, col-1, matched_product)
            changes_made += 1
            puts "Applied confirmed match: #{matched_product}"
          elsif matched_product.nil?
            unmatched_items.add(original_value)
            puts "Applied previous decision: keeping original value"
          end
          next
        end
        
        matched_product = find_closest_match(original_value, product_list)
        confirmed_matches[original_value] = matched_product
        
        if matched_product
          if cell_value != matched_product
            output_worksheet.write(row-1, col-1, matched_product)
            changes_made += 1
            puts "Changed to: #{matched_product}"
          end
        else
          unmatched_items.add(original_value)
          puts "Keeping original value"
        end
      end
    end
  end

  output_workbook.close

  puts "\nProcessing complete:"
  puts "Total cells processed: #{total_processed}"
  puts "Changes made: #{changes_made}"
  
  if !unmatched_items.empty?
    puts "\nUnmatched items:"
    unmatched_items.sort.each { |item| puts "- #{item}" }
    
    handle_unmatched_items(unmatched_items, product_list_file)
  end
  
  puts "\nResults saved to: #{output_file}"
end

def main
  default_directory = Dir.pwd
  
  product_list_file, product_list = get_product_list_file(default_directory)
  
  if product_list.empty?
    puts "Error: Product list is empty. Please check your input file"
    return
  end
  
  input_excel_file = get_input_file(default_directory)
  columns_to_clean = get_columns_to_clean
  output_excel_file = get_output_file_details(default_directory)
  
  clean_product_names(input_excel_file, columns_to_clean, output_excel_file, product_list_file, product_list)
end

if __FILE__ == $0
  begin
    main
  rescue Interrupt
    puts "\nProgram terminated by user."
  rescue StandardError => e
    puts "\nAn error occurred: #{e.message}"
  end
end