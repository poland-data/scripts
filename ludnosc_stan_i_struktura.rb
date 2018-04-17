# This is a script used to extract demographic data from
# GUS data for voivodeships, powiats and gminas
# Source files found here: http://stat.gov.pl/obszary-tematyczne/ludnosc/ludnosc/ludnosc-stan-i-struktura-w-przekroju-terytorialnym-stan-w-dniu-30-06-2017-r-,6,22.html

require 'rubyXL'
require 'json'

workbook = workbook = RubyXL::Parser.parse("./gminy.xlsx")
powiats = []
gminas = []
workbook.worksheets.each do |sheet|
    in_miasta_na_prawach_powiatu = false

    sheet.each do |row|
        next unless row
        if row[0] && row[0].value =~ /Miasta na prawach powiatu/
            in_miasta_na_prawach_powiatu = true
        end
        
        next unless row[1]
        next if in_miasta_na_prawach_powiatu && row[0].value.split('').first == ' '

        teryt = row[1].value
        next unless teryt =~ /\A\d+\z/

        target = teryt.length == 4 || in_miasta_na_prawach_powiatu ? powiats : gminas

        id = row[1].value
        name = row[0].value

        if target == powiats
            name = name.sub('Powiat ', '')
        else
            type = case name
            when /gm\.w\./ then :rural
            when /gm\. m-w\./ then 'urban-rural'.to_sym
            when /m\./ then :urban
            end

            name = name.sub('gm.w. ', '').sub('gm. m-w. ', '')
        end

        population = {}
        population[:total] = row[2].value
        population[:men] = row[3].value.to_i
        population[:women] = row[4].value.to_i
        population[:urban] = row[5].value.to_i
        population[:rural] = row[8].value.to_i

        record = {
            id: id,
            name: name,
            population: population
        }

        if target == gminas
            record[:type] = type
        else
            record[:city] = in_miasta_na_prawach_powiatu
        end

        target << record
    end
end

File.open("gminas.json", "w") do |f|
    f.puts JSON.dump(gminas)
end

File.open("powiats.json", "w") do |f|
    f.puts JSON.dump(powiats)
end
