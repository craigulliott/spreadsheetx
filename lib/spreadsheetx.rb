# zipruby is nice as it can modify an existing zip file, perfect for our usecase 
require 'zipruby'
# we use this because it comes with ruby
require 'rexml/document'
# for copying files
require 'fileutils'
# 
require 'spreadsheetx/workbook'
require 'spreadsheetx/worksheet'

module SpreadsheetX
  
  class << self

    def open(path)
      SpreadsheetX::Workbook.new(path)
    end
  
  end
  
end