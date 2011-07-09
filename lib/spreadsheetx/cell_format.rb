module SpreadsheetX

  # this object represents an existing cell format in the workbook
  class CellFormat

    attr_reader :id
    attr_reader :format
    
    def initialize(id, format)
      @id = id
      @format = format
    end
    
    def to_s
      id.to_s
    end
    
  end
end
