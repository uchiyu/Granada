require 'rexml/document'
require 'csv'
require 'json'

class Document

  def initialize()
    @file_name = ''
    @doc_num = 0
    @title = ''
    @student_num = ''
    @category = ''
    @timestamp = ''
  end

  def extract_document( text, file )
    arr = text.to_s.split(',') # テキストの結合

    @file_name = file

    # ファイル名からの情報抽出
    # 学籍番号
    @student_num = file.match(/\d{2}(T|G|t|g)\d{3}/i).to_s.strip

    # ナンバリング(n月版かを6桁の数字で)
    @doc_num = file.match(/\d{6}/).to_s.strip

    # テキストボックスからの情報抽出
    arr.each do |text_block|
      # 情報のマッチング
      # 学籍番号を抽出
      @student_num = text_block.match(/\d{2}(T|G|t|g)\d{3}\t/).to_s.strip if @student_num.strip == ''

      # ナンバリング
      @doc_num = text_block.match(/\d{4}年\d{2}月版/).to_s.delete('年').delete('月版').strip if @doc_num.strip == 0

      # タイトル
      @title = text_block.match(/^(?!.*(月例発表|\d{4}年\d{2}月版|香川大学([ ]|)工学部|stmail)).+{8,}$/).to_s.strip if @title.strip == ''

      # カテゴリ
      @category = text_block.match(/自己紹介|授業発表|制作実習|体験総括|先行研究|企業研修|技術検討|学会発表|作業実践|就職活動|中間発表|文献調査|卒論審査|学生総括/).to_s.strip if @category.strip == ''

    end

    @student_num = 's' + @student_num.to_s if @student_num.strip != ''
  end

  # pptxファイルを解凍して、一枚目のスライドの内容を抽出
  # 内容は,繋ぎで返却
  def self.slide_info( directory, file )
    %x( cp #{directory}#{file} #{file} )

    # スライドの1枚目を解凍
    %x( unzip #{file} ppt/slides/slide1.xml )

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
    FileUtils.rm(file)
    FileUtils.rm_r("./ppt")

    return text.delete('\"').delete('[').delete(']')
  end


  # timestampを格納
  def set_timestamp( directory, file )
    @timestamp = File.mtime( directory + file ).strftime("%Y-%m-%d")
  end

  def to_json(*a)
    {
      file_name: @file_name,
      doc_num: @doc_num,
      title: @title,
      student_num: @student_num,
      category: @category,
      timestamp: @timestamp
    }.to_json(*a)
  end
end
