require File.expand_path("../../jar/poi.jar", __FILE__)
require File.expand_path("../../jar/commons-io.jar", __FILE__)
$CLASSPATH << "./lib"

class ExcelCreator
  include_class org.apache.poi.hssf.usermodel.HSSFWorkbook
  include_class java.io.FileInputStream
  include_class java.io.FileOutputStream

  def create(template, data)
    fs = FileInputStream.new File.expand_path("../../templates/#{template}.xls", __FILE__)
    @wb = HSSFWorkbook.new(fs)
    sheet = @wb.get_sheet_at 0
    cell = get_cell_location(sheet, "@data")
    if cell
      cell.set_cell_value(data)
    end
    spit @wb
  end

  def dump_data(data)
    
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
        res = cell if cell.get_string_cell_value == str
      end
    end
    res
  end

end
