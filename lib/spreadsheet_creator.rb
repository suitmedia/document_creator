# Reference:
# http://spreadsheet.rubyforge.org/GUIDE_txt.html
require 'spreadsheet'
require 'pp'

Spreadsheet.client_encoding = 'UTF-8'

class SpreadsheetCreator
  TempDir = "tmp/templates/"
  PlaceholderPrefix = "PLACEHOLDER_"

  def initialize(data)
    #@data = JSON.parse data

    @template_source_path = "/home/dimas/Projects/BIIU/lib/document_creator/templates/obligasi-sun.xls"
    @data = {
            :string => {
            :TANGGAL_CETAK => "1 Januari 2011", 
            :TANGGAL_TRANSAKSI => "2 Januari 2011" },
            :row => {
            :CONTENT =>	[[1,2,3],
                    [4,5,6],
                    [7,8,9]]}}
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
    unlink_template_file
  end

  def write_to_tempfile
    @tempfile = Tempfile.new(rand.to_s)
    @worksheet.workbook.write(@tempfile)
  end

  def replace_matched_data
    (0..@worksheet.row_count).each do |row|
      (0..@worksheet.row(row).count).each do |col|
        replace_string_data(row, col)
        replace_row_data(row, col)
      end
    end
  end

  def replace_string_data(row, col)
  @data[:string].each do |k,v|
    key = "#{PlaceholderPrefix}#{k}"
      if @worksheet.cell(row,col) && @worksheet.cell(row,col).to_s.include?(key)
        @worksheet.cell(row, col).gsub!(key, v)
      end
    end
  end

  def replace_row_data(row, col)
  @data[:row].each do |k,v|
    key = "#{PlaceholderPrefix}#{k}"
    if @worksheet.cell(row,col) && @worksheet.cell(row,col).to_s.include?(key)
      @data[:row].each_with_index do |record, i|
        @worksheet.row(row).concat(record)
      end
    end
  end
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

  def unlink_template_file
    FileUtils.rm @template_path
  end

  def write_old
    @workbook = Spreadsheet::Workbook.new
    sheet = @workbook.create_worksheet

    content, header = @data["content"], @data["header"]
    sheet.row(0).concat(header)
    content.each_with_index do |row, i|	
      sheet.row(i+1).concat(row)
    end
  end

end
