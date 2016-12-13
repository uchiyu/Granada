require 'sinatra'
require 'sinatra/reloader' if development?
require 'json'
require './document'

set :bind, '133.92.165.48'

get '/documents' do
  unless params[:from] == nil
    from = params[:from]
    from = Date::strptime(from, '%Y-%m-%d') unless from == ''
  else
    from = nil
  end

  directories = []
  files = []
  find_file(directories, files, from)

  documents = []
  files.zip(directories).each do |file, directory| # zip で配列の合成
    document = Document.new()
    text = Document::slide_info(directory, file)
    document.set_timestamp(directory, file)
    document.extract_document(text, file)
    documents.push(document)
  end
  documents.to_json()
end

# チェックを行うpptxフィアルの探索
# filesには各ファイル名、directoriesには各ディレクトリ名を格納
def find_file( directories, files, from )
  # 再帰的にディレクトリを調査
  Dir.glob('./../Labo/**/*').each do |path|
    if /^(?!(.*)\/s\d{2}(T|G|t|g)\d{3}\/\d{6}\/(.*)\/(.*)).*.pptx/ =~ path # /年月/*.pptx の場合のみ
      file = path.to_s.slice!(/[^\/]*$/) # ファイル名の抽出
      directory = path.delete(path.to_s.slice!(/[^\/]*$/)) # ディレクトリ名の抽出

      next if file.match(/^(~|_).*$/)
      unless from == nil
        timestamp = File.mtime( directory + file ).strftime("%Y-%m-%d")
        next if from > Date::strptime(timestamp, '%Y-%m-%d')
      end

      #ファイル名とディレクトリ名の学籍番号が一致するならデータを格納
      if file.match(/s\d{2}(T|G|t|g)\d{3}/).to_s == directory.match(/s\d{2}(T|G|t|g)\d{3}/).to_s
        directories.push( directory )
        files.push( file )
      end
    end
  end
end

