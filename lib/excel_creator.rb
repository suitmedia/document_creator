require File.expand_path("../../jar/poi.jar", __FILE__)
require 'json'

class ExcelCreator
  include_class org.apache.poi.hssf.usermodel.HSSFWorkbook
  include_class java.io.FileInputStream
  include_class java.io.FileOutputStream

  def initialize(base)
    @base = base
  end

  def create(template, data)
    fs = FileInputStream.new File.expand_path("../../templates/#{template}", __FILE__)
    @wb = HSSFWorkbook.new(fs)
    sheet = @wb.get_sheet_at 0
    hash = JSON.parse(data)
    hash.each do |key, value|
      if cell = get_cell_location(sheet, "{#{key}}")
        dump_data(cell, value, true)
      elsif cell = get_cell_location(sheet, "[#{key}]")
        dump_data(cell, value, false)
      end
    end
    spit
  end

  def dump_data(cell, data, shift = true)
    current_cell = cell
    column_idx = cell.get_column_index
    current_row = cell.get_row
    if data.is_a? Array
      data.each_with_index do |row_data, index|
        if row_data.is_a? Array
          row_data.each do |cell_data|
            current_cell.set_cell_value(cell_data)
            current_cell = next_cell(current_cell)
          end
        else
          current_cell.set_cell_value(row_data)
        end
        if index < data.size - 1
          current_row = next_row(current_row, shift)
          current_cell = current_row.get_cell(column_idx) || current_row.create_cell(column_idx)
        end
      end
    else
      cell.set_cell_value(data)
    end
  end

  def next_cell(cell)
    row = cell.get_row
    idx = cell.get_column_index
    row.get_cell(idx + 1) || row.create_cell(idx + 1)
  end

  def next_row(row, shift)
    sheet = row.get_sheet
    row_num = row.get_row_num
    row = sheet.get_row(row_num + 1) || sheet.create_row(row_num + 1)
    if shift
      sheet.shift_rows(row_num + 1, sheet.get_last_row_num, 1)
      row = sheet.get_row(row_num + 1)
    end
    row
  end

  def spit
    tmpfile = rand.to_s
    writer = FileOutputStream.new(tmpfile)
    @wb.write(writer)
    writer.close 
    body = IO.read tmpfile
    File.delete tmpfile
    body
  end

  def get_cell_location(sheet, str)
    res = nil
    sheet.rowIterator.each do |row|
      row.cellIterator.find do |cell|
        begin
        res = cell if cell.get_string_cell_value == str
        rescue
        end
      end
    end
    res
  end

end
