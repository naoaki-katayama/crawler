# -*- coding: utf-8 -*-

require 'anemone'

# 1.upto 2 do |i|
  # urls.push "https://www.airbnb.jp/rooms/#{15772118}"
# end


0.upto 2 do |i|
  urls = ["https://www.airbnb.jp/rooms/#{i + 15772118}"]

  Anemone.crawl(urls, depth_limit: 0) do |anemone|
    anemone.on_every_page do |page|
      if page.body&.empty? == false
        puts page.url
      end

      # puts "aaa" * 100
      # p page.body.length
      # # puts page.body&.empty?
    end
  end
  sleep 0.5
end

# end
#
# Anemone.crawl(URL) do |anemone|
#   anemone.on_every_page do |page|
#     puts page.url
#   end
#
# end
