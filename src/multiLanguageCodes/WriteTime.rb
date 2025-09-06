# frozen_string_literal: true

require 'time'

# WriteTime module for writing timestamped logs.
module WriteTime
  # Appends the current timestamp to a specified file.
  #
  # @param filename [String] The name of the file to write to.
  def self.to_file(filename)
    begin
      nowtime = Time.now.strftime('%Y/%m/%d %H:%M:%S')
      log_message = "This program is written in Ruby.\nCurrent Time = #{nowtime}"

      File.open(filename, 'a') do |f|
        f.puts log_message
      end
      puts "Successfully appended log to '#{filename}'."
    rescue StandardError => e
      puts "An error occurred while writing to the file: #{e.message}"
    end
  end
end

if __FILE__ == $PROGRAM_NAME
  # Main execution block when the script is run directly.
  
  # Define the log filename.
  filename = 'MYFILE.txt'.freeze

  # Use the module to write the current time to the file.
  WriteTime.to_file(filename)
end
