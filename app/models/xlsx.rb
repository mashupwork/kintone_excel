class Xlsx
  def self.file_names
    path = Rails.root.join('tmp', 'xlsx')
    Dir.open(path) do |dir|
      dir.each do |f|
        next if f.match(/^~/)
        next unless f.match(/xlsx$/)
        puts f
      end
    end
  end

  def file filename, template
    workbook = RubyXL::Parser.parse(Rails.root.join('tmp', 'xlsx', template))
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
    path
  end

end
