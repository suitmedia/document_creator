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

  get '/create/pdf/:template' do
    content_type 'application/pdf'
    attachment
    PDFCreator.new(self).create params[:template], params[:data]
  end
end
