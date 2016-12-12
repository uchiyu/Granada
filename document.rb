require 'rexml/document'
require 'csv'

class Document

  def init
    @file_name = ''
    @doc_num = Integer.new
    @title = ''
    @date = ''
    @student_num = ''
    @category = ''
  end

  def extract_document( text, file )
    arr = text.to_s.split(',') # テキストの結合

    @file_name = file

    # ファイル名からの情報抽出
    # 学籍番号
    @student_num = file.match(/\d{2}[t|g]\d{3}/i).to_s.strip

    # ナンバリング(n月版かを6桁の数字で)
    @doc_num = file.match(/\d{6}/).to_s.strip

    # テキストボックスからの情報抽出
    arr.each do |text_block|
      # 情報のマッチング
      # 学籍番号を抽出
      @student_num = text_block.match(/\d{2}[t|g]\d{3}\t/i).to_s.strip unless @student_num.strip == ''

      # ナンバリング
      @doc_num = text_block.match(/\d{4}年\d{2}月版/).to_s.delete('年').delete('月版').strip unless @doc_num.strip == ''

      # タイトル
      @title = text_block.match(/^(?!.*(月例発表|\d{4}年\d{2}月版|香川大学([ ]|)工学部|stmail)).+{8,}$/).to_s.strip unless @title.strip == ''

      # カテゴリ
      @category = text_block.match(/自己紹介|授業発表|制作実習|体験総括|先行研究|企業研修|技術検討|学会発表|作業実践|就職活動|中間発表|文献調査|卒論審査|学生総括/).to_s.strip unless @category.strip == ''
    end
  end
end
