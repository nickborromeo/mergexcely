require 'spreadsheet'

puts "Starting the script..."

support_report = ARGV[0]

s_report = Spreadsheet.open(support_report)

customer_stats = s_report.worksheet(0)
mc_cases = s_report.worksheet(1)

merge_book = Spreadsheet::Workbook.new
merge_sheet = merge_book.create_worksheet #=> "Name of Sheet"
created_time = Time.new

puts "Created new book marked_for_closure_cases_#{created_time}.xls"

#Generate the Headers
merge_sheet.row(0).push("RespondentID")
merge_sheet.row(0).push("CollectorID")
merge_sheet.row(0).push("StartDate")
merge_sheet.row(0).push("EndDate")
merge_sheet.row(0).push("IP Address")
merge_sheet.row(0).push("Email Address")
merge_sheet.row(0).push("First Name")
merge_sheet.row(0).push("Last Name")
merge_sheet.row(0).push("Custom Data")
merge_sheet.row(0).push("Initial response time")
merge_sheet.row(0).push("Handling of support case")
merge_sheet.row(0).push("Resolution of support case")
merge_sheet.row(0).push("Overall satisfaction with support case")
merge_sheet.row(0).push("Product quality")
merge_sheet.row(0).push("Product features")
merge_sheet.row(0).push("Product usability")
merge_sheet.row(0).push("Overall satisfaction with product")
merge_sheet.row(0).push("Open-Ended Response")
merge_sheet.row(0).push("Status")
merge_sheet.row(0).push("Case Owner")

merge_book.write "./developer_cases/marked_for_closure_cases_#{created_time}.xls"

puts "Finished generating the headers"

# Merge the content of the two files based on the case number

puts "Start adding data to the new file"

count = 1 #start with row after the headers

customer_stats.each do |cssc_row|
  
  merge_sheet.row(count).push(cssc_row[0])
  merge_sheet.row(count).push(cssc_row[1])
  merge_sheet.row(count).push(cssc_row[2])
  merge_sheet.row(count).push(cssc_row[3])
  merge_sheet.row(count).push(cssc_row[4])
  merge_sheet.row(count).push(cssc_row[5])
  merge_sheet.row(count).push(cssc_row[6])
  merge_sheet.row(count).push(cssc_row[7])
  merge_sheet.row(count).push(cssc_row[8])
  merge_sheet.row(count).push(cssc_row[9])
  merge_sheet.row(count).push(cssc_row[10])
  merge_sheet.row(count).push(cssc_row[11])
  merge_sheet.row(count).push(cssc_row[12])
  merge_sheet.row(count).push(cssc_row[13])
  merge_sheet.row(count).push(cssc_row[14])
  merge_sheet.row(count).push(cssc_row[15])
  merge_sheet.row(count).push(cssc_row[16])
  merge_sheet.row(count).push(cssc_row[17])
  
  mc_cases.each do |sc_row|
    if (cssc_row[8] == sc_row[4])
      merge_sheet.row(count).push(sc_row[5])
      merge_sheet.row(count).push(sc_row[6])
      puts "Merging matched entry"
      break
    end
  end
  
  count += 1
  puts "Unmodified row"
  merge_book.write "./developer_cases/marked_for_closure_cases_#{created_time}.xls"
end

puts "Saving the workbook"
#merge_book.write "./merged_support_cases_#{Time.new}.xls"
puts "Book Saved"
