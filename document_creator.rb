class DocumentCreator < Sinatra::Base
  get '/' do
    'Document Creator'
  end

  get '/create/xls/:template' do
    ExcelCreator.new.create params[:template], params[:data]
  end
end
