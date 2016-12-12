#config: utf-8

require "rexml/document"
require "csv"

# pptxファイルを解凍して、一枚目のスライドの内容を抽出
def extract_info( directory, file )

  # スライドの1枚目を解凍
  %x( unzip #{directory}#{file} ppt/slides/slide1.xml )

  # xmlからスライド内容を抽出
  text = ''
  doc = REXML::Document.new(File.new('ppt/slides/slide1.xml'))
  doc = doc.to_s.scan(/<p:txBody>(.*?)<\/p:txBody>/)
  doc.each do |doc_elem|
    arr = doc_elem.to_s.scan(/<a:t>(.*?)<\/a:t>/)
    arr.each do |str_elem|
      text << str_elem.to_s
    end
    text << ','
  end

  # スライド解凍時に生成されたものを削除
  %x(rm -rf _rels ppt)

  return text.delete('\"').delete('[').delete(']')
end

# 登録情報の抽出
def trim_text( text, file )

  write = Array.new(4) # ファイル名, 学籍番号, ナンバリング, タイトルを格納
  arr = text.to_s.split(',') # テキストの結合

  write[0] = file

  # ファイル名からの情報抽出
  # 学籍番号
  tmp = file.match(/\d{2}[T|G|t|g]\d{3}/).to_s.strip
  if tmp != ''
    write[1] = tmp
  end

  # ナンバリング(何月版かを6桁の数字で)
  tmp = file.match(/\d{6}/).to_s.strip
  if tmp != '' || tmp != ' '
    write[2] = tmp
  end

  # テキストボックスからの情報抽出
  arr.each do |text_block|
    # 情報のマッチング
    # 学籍番号を抽出
    if write[1] == ''
      tmp = text_block.match(/\d{2}[T|G|t|g]\d{3}\t/).to_s.strip
      if ( person[0] != '' )
        write[1] = tmp
      end
    end

    # ナンバリング
    if write[2] == ''
      tmp = text_block.match(/\d{4}年\d{2}月版/).to_s.delete('年').delete('月版').strip
      if ( tmp != '' )
        write[2] = tmp
      end
    end

    # タイトル
    tmp = text_block.match(/^(?!.*(月例発表|\d{4}年\d{2}月版|香川大学([ ]|)工学部|stmail)).+{8,}$/).to_s.strip
    if tmp != ''
      write[3] = tmp
    end
  end

  # データの書き込み
  path = 'sample.csv'
  CSV.open(path, "a") do |csv|
    csv << write
  end
end

# チェックを行うpptxフィアルの探索
# filesには各ファイル名、directoriesには各ディレクトリ名を格納
def find_file( directories, files )
  # 再帰的にディレクトリを調査
  Dir.glob('./**/*').each do |path|
    if /^(?!(.*)\/s\d{2}[T|G|t|g]\d{3}\/\d{6}\/(.*)\/(.*)).*.pptx/ =~ path # /年月/*.pptx の場合のみ
      file = path.to_s.slice!(/[^\/]*$/) # ファイル名の抽出
      directly = path.delete(path.to_s.slice!(/[^\/]*$/)) # ディレクトリ名の抽出

      if file.match(/^[~$|_].*$/)
        next
      end

      #ファイル名とディレクトリ名の学籍番号が一致するならデータを格納
      if file.match(/s\d{2}[T|G|t|g]\d{3}/).to_s == directly.match(/s\d{2}[T|G|t|g]\d{3}/).to_s
        directories.push( directly )
        files.push( file )
      end
    end
  end
end

directories = Array.new
files = Array.new
find_file( directories, files )
files.zip(directories).each do |file, directory| # zip で配列の合成
  string = extract_info(directory, file)
  trim_text( string, file )
end
