require './destroyer.rb'

task :default => [:make]

task :make do
  counter = 0
  Dir.glob("data_chmi/SCE/*.{xls, xlsx}") do |file|
    Destroyer.new(file).destroy_world
    counter += 1
  end

  puts "#{counter} files was cleaned"
end
