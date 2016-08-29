#!/usr/bin/env ruby
require 'spreadsheet'
require 'date'

# http://spreadsheet.rubyforge.org/files/GUIDE_txt.html

# TODO
# 1. nevalidni dny v mesici DONE
# 2. kontrolu prazdnych hodnot v row[4] DONE
# 3. pustit na vsechny soubory ./SCE/
# 4. pocet sloupcu values pri inicializaci

class Destroyer

  def initialize(path)
    Spreadsheet.client_encoding = 'UTF-8'

    @path = path
    @document = Spreadsheet::open @path
    @empty_values = 0

    initialize_new_document
  end

  def initialize_new_document
    @new_document = Spreadsheet::Workbook.new
    @new_sheet = @new_document.create_worksheet(name: 'opraveno')
    
    # TODO remove first line from iteration
    # @new_sheet.row(0).concat @document.worksheet(0).row(0)
  end

  def destroy_world
    new_index = 0
    @document.worksheet(0).each do |row|
      if valid_row? row
        @new_sheet.row(new_index).concat row

        print_empty_value(row, new_index)
        new_index += 1
      end
    end

    save
  end

  def print_empty_value(row, new_index)
    if !row[4]
      id = row[0]
      puts "Empty value on #{id} at #{new_index}"
      @empty_values += 1
    end
  end

  def valid_row?(row)
    return (row[3] || row[4]) && valid_day?(row)
  end

  def valid_day?(row)
    year = row[1]
    month = row[2]
    day = row[3].delete('Val')
    begin
      Date.parse("#{year}#{month}#{day}")
      return true
    rescue ArgumentError
      return false
    end
  end

  def save
    new_path = @path.delete('.xls') << '_prvni_oprava.xls'
    puts @new_document.write new_path
    puts "There are #{@empty_values} empty values"
  end

end
