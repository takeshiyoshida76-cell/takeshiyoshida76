# frozen_string_literal: true

require 'time'

FILENAME = 'MYFILE.txt'.freeze

def write_time_program
  # Get System DateTime & Edit Message
  nowtime = Time.now.strftime('%Y/%m/%d %H:%M:%S')

  # Open File
  File.open(FILENAME, 'a') do |f|
    f.puts 'This program is written in Ruby.'
    f.puts "Current Time = #{nowtime}"
  end
rescue StandardError => e
  puts "ファイルへの書き込み中にエラーが発生しました: #{e.message}"
end

if __FILE__ == $PROGRAM_NAME
  write_time_program
  puts "ログが'#{FILENAME}'に追記されました。"
end
