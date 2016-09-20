#!/usr/bin/env ruby
require "google_drive"
require 'colorize'
require 'optparse'

options = {}
OptionParser.new do |opts|
  opts.banner = "Usage: example.rb [options]"

  opts.on("-v", "--[no-]verbose", "Run verbosely") do |v|
    options[:verbose] = v
  end
  
  opts.on("-h", "--help", "Prints this help") do
    puts opts
    exit
  end
        
end.parse!


class Document
  @title = nil
  @modified = nil
  @languages = nil
  @terms = nil
  @keywords = nil
  @warnings = nil
  
  attr_accessor :title, :modified, :languages, :warnings
  attr_reader :terms, :keywords
  
  def terms=(terms)
     @terms = terms
     @keywords = terms.map { |term| term.keyword }
   end
end

class Term

  attr_accessor :values, :keyword

  def initialize(keyword)
    @keyword = keyword
    @values = Hash.new
  end

  def is_comment?
    @keyword.downcase == '[comment]'
  end

end

class Diff
  
  attr_accessor :keyword, :language, :old_value, :new_value
  
  def initialize(keyword,language,old_val,new_val)
    @keyword = keyword
    @language = language
    @old_value = old_val
    @new_value = new_val
  end
  
  def log
    puts "#{@language.upcase}: #{keyword}"
    puts @old_value + " >> " + @new_value
    puts 
  end
end

class Warning

  attr_accessor :keyword, :language

  def initialize(keyword,language)
    @keyword = keyword
    @language = language
  end

end

class Processor
  def self.loadDocument(spreadsheet)
    
    doc = Document.new
    doc.title = spreadsheet.title
    
    doc.modified = spreadsheet.modified_time.strftime('%c')
 
    worksheet = spreadsheet.worksheets[0]
    raise 'Unable to retrieve the first worksheet from the spreadsheet. Are there any pages?' if worksheet.nil?

    # At this point we have the worksheet, so we want to store all the key / values
    first_valid_row_index = nil
    last_valid_row_index = nil

    for row in 1..worksheet.max_rows
      first_valid_row_index = row if worksheet[row, 1].downcase == '[key]'
      last_valid_row_index = row if worksheet[row, 1].downcase == '[end]'
    end

    raise IndexError, 'Invalid format: Could not find any [key] keyword in the A column of the worksheet' if first_valid_row_index.nil?
    raise IndexError, 'Invalid format: Could not find any [end] keyword in the A column of the worksheet' if last_valid_row_index.nil?
    raise IndexError, 'Invalid format: [end] must not be before [key] in the A column' if first_valid_row_index > last_valid_row_index

  
    languages = Hash.new('languages')
    default_language = nil

    for column in 2..worksheet.max_cols
      col_all = worksheet[first_valid_row_index, column]
      col_all.each_line(' ') do |col_text|
        default_language = col_text.downcase.gsub('*', '') if col_text.include? '*'
        languages.store col_text.downcase.gsub('*', ''), column unless col_text.to_s == ''
      end
    end

    abort 'There are no language columns in the worksheet' if languages.count == 0

    default_language = languages[0] if default_language.to_s == ''

    #puts "Languages detected: #{languages.keys.join(', ')} -- using #{default_language} as default."
    doc.languages = languages.keys
    
    #puts 'Building terminology in memory...'

    terms = []
    warnings = []
    first_term_row = first_valid_row_index+1
    last_term_row = last_valid_row_index-1

    for row in first_term_row..last_term_row
      key = worksheet[row, 1]
      unless (key.to_s == '' || key.to_s == '[comment]')
        term = Term.new(key)
        languages.each do |lang, column_index|
          term_text = worksheet[row, column_index]
          term.values.store lang, term_text
          if term_text == '' 
            warning = Warning.new(key,lang)
            warnings << warning
          end
        end
        terms << term
      end
    end
    
    doc.terms = terms
    doc.warnings = warnings
    
    doc
    
  end
end

def log_document_data(document) 
  puts document.title.yellow
  puts "Modified: #{document.modified}"
  puts "Languages: #{document.languages}"
  #puts "Keywords: #{document.keywords}"
end

########################

# 0. Parse Arguments

if ARGV.count < 1
  puts 'Source filename missing'.red 
  exit
end 

original_filename = ARGV[0]
new_filename = "CANDY_#{original_filename}"

# 1. Check documents are there
puts "1. Check documents are there".cyan

session = GoogleDrive::Session.from_config("config.json")
original_file = session.spreadsheet_by_title(original_filename)
new_file = session.spreadsheet_by_title(new_filename)


raise 'Original file missing' if original_file.nil?

raise 'New file missing' if new_file.nil?

puts "✔︎".light_green

# 2. Log documents data
puts "2. Log documents data".cyan

doc_original = Processor.loadDocument(original_file)
log_document_data(doc_original)

doc_new = Processor.loadDocument(new_file)
log_document_data(doc_new)
puts "✔︎".light_green

 
puts ""
puts "3. Check rows and columns".cyan
# 3. Check rows and columns
# - all columns are there
#TODO
# - new columns have been added
#TODO
# - all keys are there
raise IndexError, 'New document must be missing some keys' if doc_new.keywords.count < doc_original.keywords.count
# - display extra keys
 
if doc_new.keywords.count > doc_original.keywords.count
  puts "new keys: #{ doc_new.keywords - doc_original.keywords }".yellow
end
puts "✔︎".light_green
puts " "

# 4. Check cell per cell modifications
# puts doc_original.terms.count
puts "4. Check cell per cell modifications".cyan
diffs = []

doc_original.terms.each do |term|
  term.values.each_pair do |lang,content|
    #puts lang
    #puts content 
    keyword = term.keyword
    new_term = doc_new.terms.select { |term| term.keyword == keyword }
    new_content = new_term[0].values[lang]
    if content != new_content
      diff = Diff.new(keyword,lang,content,new_content)
      diffs << diff
    end
  end
end

if diffs.count > 0
  puts "Differences ".light_red
  diffs.each do |diff|
    diff.log
  end
else
  puts "No modifications".light_red
end

if doc_new.warnings.count > 0 
  puts "Warnings".light_red
  doc_new.warnings.each do |warning|
    puts "#{warning.language}: #{warning.keyword}"
  end
end







# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
#original = session.spreadsheet_by_title(original_filename).worksheets[0]
#candy = session.spreadsheet_by_title("Best Friends - Editable").worksheets[0]

# Gets content of A2 cell.
#p ntn[1, 1]  #==> "Tizio"

# Changes content of cells.
# Changes are not sent to the server until you call ws.save().
#ntn[2, 1] = "Tizio"
#ntn[2, 2] = "De Lupis"
#ntn.save
#p original.cells

#p original.cells.to_s

# Iterate over spreadsheets
# p "Original"
# p ntn.rows  #==> [["fuga", ""], ["foo", "bar]]
# p "Modified"
# p candy.rows  #==> [["fuga", ""], ["foo", "bar]]

# Diff.
# p "DIFF"
# (1..ntn.num_rows).each do |row|
#   (1..ntn.num_cols).each do |col|
#     original_item = ntn[row, col]
#     new_item = candy[row, col]
#     if original_item != new_item
#       puts "(#{row},#{col}): #{original_item} >> #{new_item}"
#     end
#   end
# end

# Reloads the worksheet to get changes by other clients.
#ntn.reload