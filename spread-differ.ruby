#!/usr/bin/env ruby
require "google_drive"



class Document
  @title = nil
  @modified = nil
  @languages = nil
  @terms = nil
  @keywords = nil
  attr_accessor :title, :modified, :languages 
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
    first_term_row = first_valid_row_index+1
    last_term_row = last_valid_row_index-1

    for row in first_term_row..last_term_row
      key = worksheet[row, 1]
      unless key.to_s == ''
        term = Term.new(key)
        languages.each do |lang, column_index|
          term_text = worksheet[row, column_index]
          term.values.store lang, term_text
        end
        terms << term
      end
    end
    
    doc.terms = terms
    
    doc
    
  end
end

def log_document_data(document) 
  puts "\# #{document.title}"
  puts "Modified: #{document.modified}"
  puts "Languages: #{document.languages}"
  puts "Keywords: #{document.keywords}"
end

original_filename = "Basic Words"
new_filename = "CANDY_#{original_filename}"
# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
session = GoogleDrive::Session.from_config("config.json")

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

########################


# 1. Check documents are there

# 2. Log documents data
# - file name
# - modified at

# 3. Check rows and columns
# - all columns are there
# - new columns have been added
# - all keys are there
# - display extra keys

# 4. Check cell per cell modifications

##########################

# 1. Check documents are there



original_file = session.spreadsheet_by_title(original_filename)
new_file = session.spreadsheet_by_title(new_filename)


raise 'Original file missing' if original_file.nil?

raise 'New file missing' if new_file.nil?


# 2. Log documents data

doc_original = Processor.loadDocument(original_file)
log_document_data(doc_original)

puts ""

doc_new = Processor.loadDocument(new_file)
log_document_data(doc_new)
 
puts ""
 
raise IndexError, 'New document must be missing some keys' if doc_new.keywords.count < doc_original.keywords.count
 
if doc_new.keywords.count > doc_original.keywords.count
  puts "new keys: #{ doc_new.keywords - doc_original.keywords }"
end

