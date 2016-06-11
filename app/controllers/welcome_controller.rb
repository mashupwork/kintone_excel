class WelcomeController < ApplicationController
  def index
    filename = 'test'
    workbook = RubyXL::Parser.parse(Rails.root.join('tmp', 'xlsx', 'download.xlsx'))
    worksheet = workbook[0]
    #CSV.parse(rendered_string).each_with_index do |row, i|
    #  row.each_with_index do |col, j|
    #    value = case col 
    #            when /\A\d+\z/
    #              col.to_i
    #            when /\A[\d\.]+\z/
    #              col.to_f
    #            else
    #              col 
    #            end 
    #    worksheet.add_cell(i, j, value)
    #  end 
    #end 
    FileUtils.mkdir_p(Rails.root.join('tmp', 'csv'))
    path = Rails.root.join('tmp', 'csv', "#{filename}.xlsx")
    workbook.write(path)
    send_file path
  end 
end

