# Reference:
# http://spreadsheet.rubyforge.org/GUIDE_txt.html
require 'spreadsheet'

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
  #     :CONTENT =>	[
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
    throw_out
  end

  private
  def write
    copy_template_file
    open_template_file
    replace_matched_data
    write_to_tempfile
    delete_template_file
  end

  def write_to_tempfile
    @tempfile = Tempfile.new(rand.to_s)
    @worksheet.workbook.write(@tempfile)
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
        @data["row"][k].each_with_index do |record,i|
          @worksheet.row(row+i).replace(record)
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
    @template = Spreadsheet.open(@template_path)
    @worksheet = @template.worksheets.first
  end

  def copy_template_file
    create_temp_dir_if_not_exist_yet
    template_path = "#{File.expand_path("")}/#{TempDir}xls-#{rand}.xls"
    FileUtils.cp(@template_source_path, template_path)	
    @template_path = template_path if File.exist?(template_path)
  end

  def create_temp_dir_if_not_exist_yet
    FileUtils.mkdir_p(TempDir) unless File.directory?(TempDir)
  end

  def throw_out
    @tempfile
  end

  def delete_template_file
    FileUtils.rm @template_path
  end
end
