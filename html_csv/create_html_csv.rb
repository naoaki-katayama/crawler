# -*- coding: utf-8 -*
# ライブラリの読み込み
require 'capybara'
require 'capybara/dsl'
# require 'selenium-webdriver'
require 'csv'
require 'pry'
# binding.pry
require 'capybara/poltergeist'
require "nokogiri"

# csv_data = CSV.read('setting.csv')
# urls = csv_data[0].first

# csv_data.each do |url_list|
# urls = csv_data[1].first

# カピバラの設定を操作

# Poltergeistで実行
Capybara.register_driver :poltergeist do |app|
  Capybara::Poltergeist::Driver.new(app, {
    :timeout => 120, js_errors: false})
end
Capybara.javascript_driver = :poltergeist
Capybara.current_driver = :poltergeist
# Capybara.app_host = urls
Capybara.default_max_wait_time = 5


module Crawler
  class Html_to_csv
    include Capybara::DSL

    def move
      # http://qiita.com/shizuma/items/7719172eb5e8c29a7d6e CSV使い方参考
      # csv_selector_setting_name = "csv_selector_setting" + ".csv"
      # csv_selector = CSV.open(csv_selector_setting_name,'w+')
      # p csv_selector
      # puts csv_selector
      csv_selector = []
      # csv_selector[n] = [] ----------------------------------------------------Arrayを作成
      csv_selector[0] = ["#"]
      urls_num = 0
      CSV.read('url_setting.csv').each do |urls|
        urls_num += 1
        visit(urls.first)
        clean_up_html(body,csv_selector,urls_num)
        csv_selector.each do |each_csv_selector|
          puts each_csv_selector.flatten
          sleep 1
        end
        # p csv_selector.flatten.size
      end
    end

    private
    def clean_up_html(origin_html,csv_selector,urls_num)
      # CSV.open('clean_uped_html.csv','w') do |csv_line|

        # csv_line << ["number","indent","tag","str","id","class","href","line"]

        pretty_html = CGI.pretty(origin_html,"fujin_raijin")
        line_count = 0

        input_judge_lines = true
        input_judge_js = true
        input_judge_noscript = true

        selector_path_tag = []
        # selector_path_detail = []
        selector_path_id = []
        selector_path_class = []
        selector_path_with_class = []
        csv_selector = []
        # for i in 0..12 do
        #   csv_selector[i] = []
        # end
        pretty_html.each_line do |line|
          line = line.chomp
          indent_number = line.scan(/fujin_raijin/).size
          indent_deleted_line = line.gsub(/fujin_raijin/,"").gsub(/\t/,"")
          space_deleted_line = indent_deleted_line.delete(' ')

          input_judge = true

          # 文末が"-->"であれば、"<!--"からの除外を次のlineから復帰
          if input_judge_lines == false
            input_judge = false
            if space_deleted_line.scan(/-->/).size == 1
              input_judge_lines = true
            end
          end

          # </script>であれば<script>からの除外を次のlineから復帰
          if input_judge_js == false
            input_judge = false
            if space_deleted_line.scan(/<\/script>/).size == 1
              input_judge_js = true
            end
          end

          # </noscript>であれば<noscript>からの除外を次のlineから復帰
          if input_judge_noscript == false
            input_judge = false
            if space_deleted_line.scan(/<\/noscript>/).size == 1
              input_judge_noscript = true
            end
          end

          # //の行を除外
          input_judge = false if /^\/\// =~ space_deleted_line

          #空白の行を除外
          input_judge = false if "" == space_deleted_line

          # <!--から除外 (-->まで)
          if /<!--.*/ =~ space_deleted_line
            input_judge = false
            unless /.*-->/ =~ space_deleted_line
              input_judge_lines = false
            end
          end

          # <script>から除外 (</script>まで)
          if /^<script.*/ =~ space_deleted_line
            input_judge = false
            input_judge_js = false
          end

          # <noscript>から除外 (</noscript>まで)
          if /^<noscript.*/ =~ space_deleted_line
            input_judge = false
            input_judge_noscript = false
          end

          # 結果をoutput
          if input_judge == true
            # puts indent_deleted_line
            # binding.pry
            extracted_tag = extract_tag(indent_deleted_line)
            extracted_str = indent_deleted_line unless indent_deleted_line[0] == "<"
            extracted_id = extract_tag_attribute(indent_deleted_line,"id")
            extracted_class = extract_tag_attribute(indent_deleted_line,"class")
            extracted_href = extract_tag_attribute(indent_deleted_line,"href")

            if indent_number == 0
              selector_path_tag[0] = extracted_tag
              selector_path_tag.pop(selector_path_tag.length - 1)
              if extracted_id == nil
                selector_path_id[0] = ""
              else
                selector_path_id[0] = extracted_id
              end
              selector_path_id.pop(selector_path_id.length - 1)

              if extracted_class == nil
                selector_path_class[0] = ""
              else
                selector_path_class[0] = extracted_class
              end
              selector_path_class.pop(selector_path_class.length - 1)

              # selector_path_detail[0] = extracted_tag
              # selector_path_detail.pop(selector_path_detail.length - 1)
            else
              # binding.pry
              selector_path_tag[indent_number] = extracted_tag
              selector_path_tag.pop(selector_path_tag.length - indent_number - 1)

              if extracted_id == nil
                selector_path_id[indent_number] = ""
              else
                selector_path_id[indent_number] = extracted_id
              end
              selector_path_id.pop(selector_path_id.length - indent_number - 1)

              if extracted_class == nil
                selector_path_class[indent_number] = ""
              else
                selector_path_class[indent_number] = extracted_class
              end
              selector_path_class.pop(selector_path_class.length - indent_number - 1)

              # クラス入りのセレクタを作成
              selector_path_with_class[indent_number] = extracted_tag
              #selectorようのクラスを作成（スペースをなくし、.とする。文頭のスペースは削除）
              # class_for_selector = unless extracted_class == nil
              selector_path_with_class[indent_number] = selector_path_with_class[indent_number] + selector_path_with_class[indent_number] + "." + extracted_class unless extracted_class == nil

              # selector_path_detail[indent_number] = extracted_tag
              # selector_path_detail[indent_number] = selector_path_detail[indent_number] + "#" + extracted_id unless extracted_id == nil
              # selector_path_detail[indent_number] = selector_path_detail[indent_number] + "." + extracted_class unless extracted_class == nil
              # selector_path_detail.pop(selector_path_detail.length - indent_number - 1)
            end #if indent_number == 0

            extracted_selector_path_tag = selector_path_tag.join(' ')
            # binding.pry
            if selector_path_with_class == nil
              extracted_selector_path_with_class = []
            else
              extracted_selector_path_with_class = selector_path_with_class.join(' ')
            end

            # 整理
            csr_selector_row = [urls_num,line_count,indent_number,extracted_tag,extracted_str,extracted_id,extracted_class,extracted_href,extracted_selector_path_tag,extracted_selector_path_with_class,selector_path_id,indent_deleted_line]
            p csr_selector_row
            csv_selector.push(csr_selector_row)
            # csv_selector[0].push(urls_num) #crowlingページ数
            # csv_selector[1].push(line_count)　#各ページのhtml行数
            # csv_selector[2].push(indent_number)　#タグの階層
            # csv_selector[3].push(extracted_tag)
            # csv_selector[4].push(extracted_str)
            # csv_selector[5].push(extracted_id)
            # csv_selector[6].push(extracted_class)
            # csv_selector[7].push(extracted_href)
            # csv_selector[8].push(extracted_selector_path_tag) #タグのみのセレクタ
            # csv_selector[9].push(extracted_selector_path_with_class) #クラス入りセレクタ
            # csv_selector[10].push(selector_path_class) #Class配列
            # csv_selector[11].push(selector_path_id) #ID配列
            # csv_selector[12].push(indent_deleted_line)

            # binding.pry if line_count == 19
            # csv_line << [line_count,indent_number,extracted_tag,extracted_str,extracted_id,extracted_class,extracted_href,indent_deleted_line,extracted_selector_path_tag,extracted_selector_path_detail]
            # csv.push = [line_count,indent_number,extracted_tag,extracted_str,extracted_id,extracted_class,extracted_href,indent_deleted_line]
            # puts extracted_selector_path_tag
            # puts extracted_selector_path_detail

            # p "#{line_count},#{indent_number},#{extracted_tag},#{extracted_str},#{extracted_id},#{extracted_class},#{extracted_href},#{indent_deleted_line}"
            line_count += 1
          else
            # puts indent_deleted_line
          end
        end # pretty_html.each_line
      # end # csv
    end

    # nokogiriを活用して、attributeの要素を抽出する
    def extract_tag_attribute(str,attribute)
      if !(extract_tag(str)[0] == "\/" || extract_tag(str)[0] == "!" || extract_tag(str)[0] == "$")
        if str[0] == "<"
          return Nokogiri::HTML.parse(str).css(extract_tag(str)).first[attribute]
        end
      end
    end #extract_tag_attribute

    def extract_tag(str)
      if str[0] == "<"
        if str.index(" ") == nil
          return str[1..str.size - 2]
        else
          return str[1..str.index(" ") - 1]
        end
      else
        return ""
      end
    end

  end
end

crawler = Crawler::Html_to_csv.new
crawler.move
