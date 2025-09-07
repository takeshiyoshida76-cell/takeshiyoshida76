# Ruby script for writing current time to a file (with error handling)

begin
  now_str = Time.now.strftime('%Y/%m/%d %H:%M:%S')
  output_file = 'MYFILE.txt'

  File.open(output_file, 'w') do |f|
    f.puts "This program was written in Ruby."
    f.puts "Current time is: #{now_str}"
  end

  puts "Successfully wrote current time to '#{output_file}'."
rescue StandardError => e
  warn "An error occurred while writing to the file: #{e.message}"
  exit 1
end
