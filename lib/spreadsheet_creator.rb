# Reference:
# http://spreadsheet.rubyforge.org/GUIDE_txt.html
require 'spreadsheet'
Spreadsheet.client_encoding = 'UTF-8'

class SpreadsheetCreator
	def initialize(data)
	  @data = JSON.parse data
	end

	def create
	  write
		throw_out
	end

private
	def write
	  @workbook = Spreadsheet::Workbook.new
		sheet = @workbook.create_worksheet

		content, header = @data.first, @data.last
		sheet.row(1).replace(header)
		content.each_with_index do |row, i|	
			sheet.row(i+2).push(row)
		end
	end

	def throw_out
	  tmp_file = Tempfile.new(rand.to_s)
		@workbook.write(tmp_file)
		tmp_file
	end
end
