class DocumentCreator < Sinatra::Base
  get '/' do
    'Document Creator'
  end

  post '/create/xls/:template' do
    content_type 'application/vnd.ms-excel'
    attachment
    ExcelCreator.new.create params[:template], params[:data]
  end
end
