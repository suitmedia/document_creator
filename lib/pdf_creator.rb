require 'open3'

class PDFCreator
  def initialize(base)
    @base = base
    @wkhtmltopdf = File.expand_path("../../bin/wkhtmltopdf-i386", __FILE__)
  end

  def create(template, data)
    str = @base.haml template.to_sym
    pdf_from_string(str)
  end

  def pdf_from_string(str)
    command = "#{@wkhtmltopdf} - -"
    Open3.popen3(command) do |i, o, e|
      i.write(str)
      i.close
      o.read
    end
  rescue
    "Error"
  end
end
