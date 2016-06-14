class Xlsx
  def self.file_names
    res = []
    path = Rails.root.join('tmp', 'xlsx')
    Dir.open(path) do |dir|
      dir.each do |f|
        next if f.match(/^~/)
        next unless f.match(/xlsx$/)
        res.push(f)
      end
    end
    res
  end

  def self.file filename, template
    workbook = RubyXL::Parser.parse(Rails.root.join('tmp', 'xlsx', template))
    worksheet = workbook[0]

    # https://github.com/weshatheleopard/rubyXL
    # Possible weights: hairline, thin, medium, thick
    #
    # http://www.rubydoc.info/gems/rubyXL/3.3.17/RubyXL
    # %w{ none thin medium dashed dotted thick double hair
    # mediumDashed dashDot mediumDashDot dashDotDot slantDashDot }
    5.upto(10).each do |i|
      worksheet.add_cell(i, 1, 'hoge')
      #worksheet.sheet_data[i][1].change_border(:top, 'thin') # 普通の線
      #worksheet.sheet_data[i][1].change_border(:top, 'hairline') # 線なし
      #worksheet.sheet_data[i][1].change_border(:top, 'medium') # 太線
      worksheet.sheet_data[i][1].change_border(:top, 'dotted') # 点線
    end

    FileUtils.mkdir_p(Rails.root.join('tmp', 'csv'))
    path = Rails.root.join('tmp', 'csv', "#{filename}.xlsx")
    workbook.write(path)
    path
  end
end

