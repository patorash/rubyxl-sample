require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'pry'

@workbook = RubyXL::Parser.parse('sample.xlsx')
@worksheet = @workbook[0]

@worksheet.each_with_index do |row, i|
  next if i.zero?
  row[1].change_contents(i + 1)
end
# 1.upto(@worksheet.)
# @worksheet.each.with_index(5) do |row, i|
#   row[1] = i
# end
@workbook.write("./update_sample.xlsx")
# @workbook.
# binding.pry

# @workbook
# true