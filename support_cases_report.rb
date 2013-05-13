require 'spreadsheet'

puts "Starting the script..."

support_report = ARGV[0]

s_report = Spreadsheet.open(support_report)

support_stats = s_report.worksheet(0)
support_cases = s_report.worksheet(1)

merge_book = Spreadsheet::Workbook.new
merge_sheet = merge_book.create_worksheet #=> "Name of Sheet"
created_time = Time.new

puts "Created new book merged_support_cases_#{created_time}.xls"

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
merge_sheet.row(0).push("Last Date Modified")
merge_sheet.row(0).push("Age")
merge_sheet.row(0).push("Status")
merge_sheet.row(0).push("Case Origin")
merge_sheet.row(0).push("Region")
merge_sheet.row(0).push("Created By")
merge_sheet.row(0).push("Case Record Type")
merge_sheet.row(0).push("Initial response time")
merge_sheet.row(0).push("Handling of support case")
merge_sheet.row(0).push("Resolution of support case")
merge_sheet.row(0).push("Overall satisfaction with support case")
merge_sheet.row(0).push("Product quality")
merge_sheet.row(0).push("Product features")
merge_sheet.row(0).push("Product usability")
merge_sheet.row(0).push("Overall satisfaction with product")
merge_sheet.row(0).push("Open-Ended Response")

merge_book.write "./support_cases/merged_support_cases_#{created_time}.xls"

puts "Finished generating the headers"

# Merge the content of the two files based on the case number

puts "Start adding data to the new file"

count = 1 #start with row after the headers
line_number = 0

support_cases.each do |sc_row|
  puts "Processing #{line_number += 1}"
  support_stats.each do |cssc_row|
    if (cssc_row[8] == sc_row[0])
      merge_sheet.row(count).push(sc_row[0])
      merge_sheet.row(count).push(sc_row[1])
      merge_sheet.row(count).push(sc_row[2])
      merge_sheet.row(count).push(sc_row[3])
      merge_sheet.row(count).push(sc_row[4])
      merge_sheet.row(count).push(sc_row[5])
      merge_sheet.row(count).push(sc_row[6])
      merge_sheet.row(count).push(sc_row[7])
      merge_sheet.row(count).push(sc_row[8])
      merge_sheet.row(count).push(sc_row[9])
      merge_sheet.row(count).push(sc_row[10])
      merge_sheet.row(count).push(sc_row[11])
      merge_sheet.row(count).push(sc_row[12])
      merge_sheet.row(count).push(sc_row[13])
      merge_sheet.row(count).push(sc_row[14])
      merge_sheet.row(count).push(sc_row[15])
      merge_sheet.row(count).push(sc_row[16])
      merge_sheet.row(count).push(sc_row[17])
      merge_sheet.row(count).push(cssc_row[9])
      merge_sheet.row(count).push(cssc_row[10])
      merge_sheet.row(count).push(cssc_row[11])
      merge_sheet.row(count).push(cssc_row[12])
      merge_sheet.row(count).push(cssc_row[13])
      merge_sheet.row(count).push(cssc_row[14])
      merge_sheet.row(count).push(cssc_row[15])
      merge_sheet.row(count).push(cssc_row[16])
      merge_sheet.row(count).push(cssc_row[17])
      count += 1
      puts "Merging matched entry"
      merge_book.write "./support_cases/merged_support_cases_#{created_time}.xls"
    end
  end
end

puts "Saving the workbook"
#merge_book.write "./merged_support_cases_#{Time.new}.xls"
puts "Book Saved"
