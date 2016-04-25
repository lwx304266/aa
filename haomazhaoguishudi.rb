require 'rubygems'
require 'selenium-webdriver'
require 'win32ole'


excel = WIN32OLE.new('excel.application')
excel.visible = true
workbook = excel.Workbooks.Open('C:/Users/LN/Desktop/demo/number.xlsx')
worksheet = workbook.Worksheets(1) #定位到第一个sheet
worksheet.Select
driver = Selenium::WebDriver.for :firefox
driver.get('http://www.ip138.com:8080/search.asp?')
i = 1
while worksheet.Range("A#{i}").Value !='' and !worksheet.Range("A#{i}").Value.nil?
	number = worksheet.Range("A#{i}").Value
	driver.find_element(:name, 'mobile').send_keys(number)
	driver.find_element(:class, 'bdtj').send_keys("\n")
    begin
		element = driver.find_elements(:class, 'tdc2')
		worksheet.Range("B#{i}").Value = element[1].text
		worksheet.Range("C#{i}").Value = element[2].text
	rescue
		worksheet.Range("B#{i}").Value = '号码有误!'
		11.times{driver.find_element(:name, 'mobile').send_keys("\b")}
	ensure
		i+=1
	end
end
driver.quit
workbook.Close(1)
excel.quit