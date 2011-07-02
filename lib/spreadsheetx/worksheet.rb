module SpreadsheetX

  # Workbooks are made up of N Worksheets, this class represents a specific Worksheet.
  class Worksheet

    attr_reader :sheet_id
    attr_reader :r_id
    attr_reader :name
    
    # return a Worksheet object which relates to a specific Worksheet
    def initialize(archive, sheet_id, r_id, name)
      @sheet_id = sheet_id
      @r_id = r_id
      @name = name

      # open the workbook
      archive.fopen("xl/worksheets/sheet#{@r_id}.xml") do |f| 

        # read contents of this file
        file_contents = f.read 
        #parse the XML and hold the doc
        @xml_doc = REXML::Document.new(file_contents)

      end
      
    end

    # update the value of a particular cell, if the row or cell doesnt exist in the XML, then it will be created
    def update_cell(col_number, row_number, val)
      
      cell_id = SpreadsheetX::Worksheet.cell_id(col_number, row_number)
      
      rows = @xml_doc.get_elements("worksheet/sheetData/row[@r=#{row_number}]")
      # was this row found
      if rows.empty?
        # build a new row
        row = @xml_doc.elements['worksheet'].elements['sheetData'].add_element('row', {'r' => row_number})
      else
        # x path returns an array, but we know there is only one row with this number
        row = rows.first
      end
      
      cells = row.get_elements("c[@r='#{cell_id}']")
      if cells.empty?
        cell = row.add_element('c', {'r' => cell_id})
      else
        # x path returns an array, but we know there is only one row with this number
        cell = cells.first
      end
      
      # first clear out any existing values
      cell.delete_element('*')

      # now we put the value in the cell
      if val.kind_of? String
        cell.attributes['t'] = 'inlineStr'
        cell.add_element('is').add_element('t').add_text(val)
      else
        cell.attributes['t'] = nil
        cell.add_element('v').add_text(val.to_s)
      end

    end
    
    # the number of rows containing data this sheet has 
    # NOTE: this is the count of those rows, not the length of the document
    def row_count
      count = 0
      @xml_doc.elements.each('worksheet/sheetData/row'){ count+=1 }
      count
    end

    # returns the xml representation of this worksheet
    def to_s
      @xml_doc.to_s
    end
    
    # turns a cell address into its excel name, 1,1 = A1  2,3 = C2 etc.
    def self.cell_id(col_number, row_number)
      letter = 'A'
      # some day, speed this up
      (col_number.to_i-1).times{letter = letter.succ}
      "#{letter}#{row_number}"
    end
    
  end
  
end
