module SpreadsheetX

  # This class represents an XLSX Document on disk
  class Workbook

    attr_reader :path
    attr_reader :worksheets

    # return a Workbook object which relates to an existing xlsx file on disk
    def initialize(path)
      @path = path
      Zip::Archive.open(path) do |archive|
        
        # open the workbook
        archive.fopen('xl/workbook.xml') do |f| 

          # read contents of this file
          file_contents = f.read 

          #parse the XML and build the worksheets
          @worksheets = []
          # parse the XML and hold the doc
          xml_doc = XML::Document.string(file_contents)
          # set the default namespace
          xml_doc.root.namespaces.default_prefix = 'spreadsheetml'
          
          xml_doc.find('spreadsheetml:sheets/spreadsheetml:sheet').each do |node|
            sheet_id = node['sheetId'].to_i
            r_id = node['id'].gsub('rId','').to_i
            name = node['name'].to_s
            @worksheets.push SpreadsheetX::Worksheet.new(archive, sheet_id, r_id, name)
          end
          
        end

      end
    end
    
    # saves the binary form of the complete xlsx file to a new xlsx file
    def save(destination_path)

      # copy the xlsx file to the destination
      FileUtils.cp(@path, destination_path)
      
      # replace the xlsx files with the new workbooks
      Zip::Archive.open(destination_path) do |ar|
        
        # replace with the new worksheets
        @worksheets.each do |worksheet|
          ar.replace_buffer("xl/worksheets/sheet#{worksheet.r_id}.xml", worksheet.to_s)
        end
                
      end

    end

  end
  
end

