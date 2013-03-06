require 'spreadsheet'

puts "Starting the script..."

customer_service_report = ARGV[0]

cs_report = Spreadsheet.open(customer_service_report)

customer_service_stats = cs_report.worksheet(0)
customer_service_cases = cs_report.worksheet(1)

merge_book = Spreadsheet::Workbook.new
merge_sheet = merge_book.create_worksheet #=> "Name of Sheet"
created_time = Time.new

puts "Created new book merged_customer_service_cases_#{created_time}.xls"

#Generate the Headers
merge_sheet.row(0).push("Case Number")
merge_sheet.row(0).push("Type")
merge_sheet.row(0).push("Bug ID")
merge_sheet.row(0).push("Product Family")
merge_sheet.row(0).push("Product")
merge_sheet.row(0).push("Product Version")
merge_sheet.row(0).push("Account Name")
merge_sheet.row(0).push("Contact Name")
merge_sheet.row(0).push("Subject")
merge_sheet.row(0).push("Case Owner")
merge_sheet.row(0).push("Opened Date")
merge_sheet.row(0).push("Case Last Modified Date")
merge_sheet.row(0).push("Age")
merge_sheet.row(0).push("Status")
merge_sheet.row(0).push("Case Origin")
merge_sheet.row(0).push("Region")
merge_sheet.row(0).push("Created By")
merge_sheet.row(0).push("Case Record Type")
merge_sheet.row(0).push("Response Time")
merge_sheet.row(0).push("Communication")
merge_sheet.row(0).push("Resolution")
merge_sheet.row(0).push("Overall satisfaction")
merge_sheet.row(0).push("Open-Ended Response")
merge_sheet.row(0).push("How did you contact support?")

merge_book.write "./cs_cases/merged_customer_service_cases_#{created_time}.xls"

puts "Finished generating the headers"

# Merge the content of the two files based on the case number

puts "Start adding data to the new file"

count = 1 #start with row after the headers

customer_service_stats.each do |customer_sat_row|
  customer_service_cases.each do |service_row|
    if (customer_sat_row[0] == service_row[0])
      merge_sheet.row(count).push(service_row[0])
      merge_sheet.row(count).push(service_row[1])
      merge_sheet.row(count).push(service_row[2])
      merge_sheet.row(count).push(service_row[3])
      merge_sheet.row(count).push(service_row[4])
      merge_sheet.row(count).push(service_row[5])
      merge_sheet.row(count).push(service_row[6])
      merge_sheet.row(count).push(service_row[7])
      merge_sheet.row(count).push(service_row[8])
      merge_sheet.row(count).push(service_row[9])
      merge_sheet.row(count).push(service_row[10])
      merge_sheet.row(count).push(service_row[11])
      merge_sheet.row(count).push(service_row[12])
      merge_sheet.row(count).push(service_row[13])
      merge_sheet.row(count).push(service_row[14])
      merge_sheet.row(count).push(service_row[15])
      merge_sheet.row(count).push(service_row[16])
      merge_sheet.row(count).push(service_row[17])
      merge_sheet.row(count).push(customer_sat_row[1])
      merge_sheet.row(count).push(customer_sat_row[2])
      merge_sheet.row(count).push(customer_sat_row[3])
      merge_sheet.row(count).push(customer_sat_row[4])
      merge_sheet.row(count).push(customer_sat_row[5])
      merge_sheet.row(count).push(customer_sat_row[6])
     
      count += 1
      puts "Merging matched entry"
      merge_book.write "./cs_cases/merged_customer_service_cases_#{created_time}.xls"
    end
  end
end

puts "Saving the workbook"
#merge_book.write "./merged_support_cases_#{Time.new}.xls"
puts "Book Saved"
