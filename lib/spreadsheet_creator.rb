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

		content, header = @data["content"], @data["header"]
		sheet.row(0).concat(header)
		content.each_with_index do |row, i|	
			sheet.row(i+1).concat(row)
		end
	end

	def throw_out
	  @tmp_file = Tempfile.new(rand.to_s)
		@workbook.write(@tmp_file)
		@tmp_file
	end

	def unlink_tempfile
	  @tmp_file.delete
	end
end
