# Reference:
# http://spreadsheet.rubyforge.org/GUIDE_txt.html
require 'spreadsheet'
require 'stringio'

Spreadsheet.client_encoding = 'UTF-8'

class SpreadsheetCreator
  TempDir = "tmp/templates/"
  PlaceholderPrefix = "PLACEHOLDER_"

  # Use this format for 'data' parameter:
  # @data = {
  #   :string => {
  #     :TANGGAL_CETAK => "1 Januari 2011", 
  #     :TANGGAL_TRANSAKSI => "2 Januari 2011" },
  #   :row => {
  #     :CONTENT =>  [
  #           [1,2,3],
  #           [4,5,6],
  #           [7,8,9]]
  # }}

  def initialize(data, template_source_path)
    @data = JSON.parse data
    @template_source_path = template_source_path
  end

  def create
    write
  end

  private
  def write
    open_template_file
    replace_matched_data
    write_to_io
  end

  def write_to_io
    io = StringIO.new
    @worksheet.workbook.write(io)
    io.rewind
    io.read
  end

  def replace_matched_data
    (0..@worksheet.row_count).each do |row|
      (0..@worksheet.row(row).count).each do |col|
        replace_string_data_if_found(row, col)
        replace_row_data_if_found(row, col)
      end
    end
  end

  def replace_string_data_if_found(row, col)
    @data["string"].each do |k,v|
      key = spreadsheet_key(k)
      if cell_value_included_key?(row, col, key)
        @worksheet.cell(row, col).gsub!(key, v)
      end
    end
  end

  def replace_row_data_if_found(row, col)
    @data["row"].each do |k,v|
      key = spreadsheet_key(k)
      if cell_value_included_key?(row,col,key)
        @data["row"][k].each do |records|
          records.each_with_index do |record,i|
            @worksheet.row(row+i).replace(record)
          end
        end
      end
    end
  end

  def cell_value_included_key?(row,col,key)
    @worksheet.cell(row,col) && @worksheet.cell(row,col).to_s.include?(key)
  end

  def spreadsheet_key(key)
    "#{PlaceholderPrefix}#{key}"
  end

  def open_template_file
    @template = Spreadsheet.open(@template_source_path)
    @worksheet = @template.worksheets.first
  end

end
