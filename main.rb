require 'spreadsheet'
require 'nokogiri'


path = 'C:\\Users\\Andrei_Labetka\\Desktop\\temp\\'
TRXNAME = 'TestResults.trx'
EXELNAME = 'TestResults.xls'

xml_file = path + TRXNAME
failed_tests = []
project_name = ''

xml = Nokogiri::XML(File.open xml_file) do |config|
  config.huge
end

xml.css('Results UnitTestResult[outcome="Failed"]').each do |node|
  testname = node['testName']
  failed_tests.push testname
end

mas = []
xml.css('TestCategory').each do |item|
  mas << item.css('TestCategoryItem').map{|i| i['TestCategory']}
end
project_name = mas.flatten.each_with_object(Hash.new(0)){ |m,h| h[m] += 1 }.sort_by{ |k,v| v }.select{|x| x[1]==mas.length}.map{|x| x.first}.join(' ')
xml = nil


Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet name: 'My Second Worksheet'

sheet1.row(0).concat %w{N\ # Tester Projects Environment Browser Failed\ Tests Comments Passed\ Tests Total\Tests Passed\Checks Total\ Checks Checks\ pass\ rate }

sheet1[1,3] = 'Japan'



book.write path + EXELNAME


