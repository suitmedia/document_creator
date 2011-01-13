require 'open3'
require 'json'

class PDFCreator

  def initialize(base)
    @base = base
    @wkhtmltopdf = File.expand_path("../../bin/wkhtmltopdf-i386", __FILE__)
  end

  def create(template, data)
    @base.json_to_instance_variables(data)
    str = @base.haml template.to_sym
    str_to_pdf(str)
  end

  def str_to_pdf(str)
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
