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
        # parse the XML and hold the doc
        @xml_doc = XML::Document.string(file_contents)
        # set the default namespace
        @xml_doc.root.namespaces.default_prefix = 'spreadsheetml'

      end
      
    end

    # update the value of a particular cell, if the row or cell doesnt exist in the XML, then it will be created
    def update_cell(col_number, row_number, val)
      
      cell_id = SpreadsheetX::Worksheet.cell_id(col_number, row_number)
      
      row = @xml_doc.find_first("spreadsheetml:sheetData/spreadsheetml:row[@r=#{row_number}]")
      
      # was this row found
      unless row
        # build a new row
        row = XML::Node.new('row')
        row['r'] = row_number.to_s
        # add it to the other rows
        @xml_doc.find_first('spreadsheetml:sheetData') << row
      end
      
      cell = row.find_first("spreadsheetml:c[@r='#{cell_id}']")
      # was this row found
      unless cell
        # build a new cell
        cell = XML::Node.new('c')
        cell['r'] = cell_id
        # add it to the other rows
        row << cell
      end
      
      # create the node which represents the value in the cell
      if val.kind_of? String
        
        # put the strings inline to make life easier
        cell['t'] = 'inlineStr'
        
        # the string node looke like <is><t>string</t></is>
        is = XML::Node.new('is')
        t = XML::Node.new('t')
        t.content = val
        
        cell_value = ( is << t )
        
      else
        
        # incase this was an inline string, clear out this attribute
        cell['t'] = ''
        
        cell_value = XML::Node.new('v')
        cell_value.content = val.to_s

      end
      
      # first clear out any existing values (nodes)
      cell.find('*').each{|n| n.remove! }

      # now we put the value in the cell
      cell << cell_value

    end
    
    # the number of rows containing data this sheet has 
    # NOTE: this is the count of those rows, not the length of the document
    def row_count
      count = 0
      # target the sheetData rows
      @xml_doc.find('spreadsheetml:sheetData/spreadsheetml:row').count
    end

    # returns the xml representation of this worksheet
    def to_s
      @xml_doc.to_s
    end
    
    # turns a cell address into its excel name, 1,1 = A1  2,3 = C2 etc.
    def self.cell_id(col_number, row_number)
      raise 'There is no row 0, start at 1 instead' if row_number < 1
      letter = 'A'
      # some day, speed this up
      (col_number.to_i-1).times{letter = letter.succ}
      "#{letter}#{row_number}"
    end
    
  end
  
end
