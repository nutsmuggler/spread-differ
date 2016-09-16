require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
session = GoogleDrive::Session.from_config("config.json")

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
ntn = session.spreadsheet_by_title("Best Friends").worksheets[0]
candy = session.spreadsheet_by_title("Best Friends - Editable").worksheets[0]

# Gets content of A2 cell.
#p ntn[1, 1]  #==> "Tizio"

# Changes content of cells.
# Changes are not sent to the server until you call ws.save().
#ntn[2, 1] = "Tizio"
#ntn[2, 2] = "De Lupis"
#ntn.save

# Iterate over spreadsheets
p "Original"
p ntn.rows  #==> [["fuga", ""], ["foo", "bar]]
p "Modified"
p candy.rows  #==> [["fuga", ""], ["foo", "bar]]

# Diff.
p "DIFF"
(1..ntn.num_rows).each do |row|
  (1..ntn.num_cols).each do |col|
    original_item = ntn[row, col]
    new_item = candy[row, col]
    if original_item != new_item
      puts "(#{row},#{col}): #{original_item} >> #{new_item}"
    end
  end
end


# Reloads the worksheet to get changes by other clients.
ntn.reload