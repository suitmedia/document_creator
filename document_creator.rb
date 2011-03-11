class DocumentCreator < Sinatra::Base

  set :views, File.dirname(__FILE__) + '/templates'

  get '/' do
    'Document Creator'
  end

  post '/create/xls/:template' do
    content_type 'application/vnd.ms-excel'
    attachment
    ExcelCreator.new(self).create params[:template], params[:data]
  end

  post '/create/spreadsheet' do
    content_type 'application/vnd.ms-excel'
    attachment
    SpreadsheetCreator.new(params[:data], params[:template]).create
  end

  post '/create/pdf/:template' do
    content_type 'application/pdf'
    attachment
    PDFCreator.new(self).create params[:template], params[:data]
  end

  def json_to_instance_variables(json)
    hash = JSON.parse json
    hash.each do |k, v|
      self.instance_variable_set(k, v)
    end
  end
end
