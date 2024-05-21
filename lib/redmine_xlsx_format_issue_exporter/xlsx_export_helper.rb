require 'write_xlsx'

module RedmineXlsxFormatIssueExporter
  module XlsxExportHelper

    def query_to_xlsx(items, query, options={})
  columns = query.columns
  extra_columns = ["nemkell", "rag_oszlop", "hat_oszlop", "elo_oszlop", "Teljes ut hossza", "Utido", "Munkaido"]
  extra_columns_size = extra_columns.size
  columns = extra_columns + columns

  stream = StringIO.new('')
  workbook = WriteXLSX.new(stream)
  worksheet = workbook.add_worksheet

  worksheet.freeze_panes(1, 1)  # Freeze header row and # column.

  columns_width = []

  # Write the header row with extra columns
  write_header_row(workbook, worksheet, columns, columns_width)

  # Write item rows with extra columns
  row_number = 2
  write_item_rows(workbook, worksheet, columns, items, columns_width)

  columns.size.times do |index|
    worksheet.set_column(index + 7, index + 7, columns_width[index])
  end

  workbook.close

  stream.string
end


    def write_header_row(workbook, worksheet, columns, columns_width)
      header_format = create_header_format(workbook)
      columns.each_with_index do |c, index|
        if c.class.name == 'String'
            value = c
        else
            value = c.caption.to_s
        end

        worksheet.write(0, index, value, header_format)
        columns_width << get_column_width(value)
      end
    end

    def write_item_rows(workbook, worksheet, columns, items, columns_width)
      hyperlink_format = create_hyperlink_format(workbook)
      cell_format = create_cell_format(workbook)
      
      # Skip the first 7 columns
      custom_columns = columns[0..6]
      original_columns = columns[7..-1] || []
      

      row_number = 1
      items.each_with_index do |item, item_index|
        row_number += 1
        custom_columns.each_with_index do |c, column_index|
          
          custom_data = case column_index
          when 0
            
            "=IF(AND(B#{row_number}=0,C#{row_number}=0,D#{row_number}=0,E#{row_number}=0,F#{row_number}=0,G#{row_number}=0),1,0)"
          when 1
            "=IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Ragasztó eltáv. oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW())=\"\",IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matrica típusa\",$1:$1,0),4),\"1\",\"\") & ROW())=\"bármelyik ragasztóeltávolítással\",INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matricázás oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()),),INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Ragasztó eltáv. oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()))"
          when 2
            "=IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hátfal oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW())=\"\",IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matrica típusa\",$1:$1,0),4),\"1\",\"\") & ROW())=\"hát ragasztóeltávolítás nélkül\",INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matricázás oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()),),INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hátfal oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()))"
          when 3
            "=IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Elő- és oldalfal oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW())=\"\",IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matrica típusa\",$1:$1,0),4),\"1\",\"\") & ROW())=\"elő és oldal ragasztóeltávolítás nélkül\",INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Matricázás oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()),),INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Elő- és oldalfal oszlopszám\",$1:$1,0),4),\"1\",\"\") & ROW()))"
          when 4
            "=INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Megérkezés km óraállás\",$1:$1,0),4),\"1\",\"\") & ROW())-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Induló km óraállás\",$1:$1,0),4),\"1\",\"\") & ROW())+IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hazaérkezés km óraállás\",$1:$1,0),4),\"1\",\"\") & ROW())>0,INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hazaérkezés km óraállás\",$1:$1,0),4),\"1\",\"\") & ROW())-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Induló km óraállás\",$1:$1,0),4),\"1\",\"\") & ROW()))-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Kitérő hossza\",$1:$1,0),4),\"1\",\"\") & ROW())"
          when 5
            "=ROUNDUP((INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Megérkezés időpontja\",$1:$1,0),4),\"1\",\"\") & ROW())-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Elindulás időpontja\",$1:$1,0),4),\"1\",\"\") & ROW()))*24*60-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Kitérő ideje\",$1:$1,0),4),\"1\",\"\") & ROW())+IF(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hazaérkezés időpontja\",$1:$1,0),4),\"1\",\"\") & ROW())<>\"\",(INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Hazaérkezés időpontja\",$1:$1,0),4),\"1\",\"\") & ROW())-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Munkavégzés befejezése\",$1:$1,0),4),\"1\",\"\") & ROW()))*24*60),0)"
          when 6
            "=ROUNDUP((INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Munkavégzés befejezése\",$1:$1,0),4),\"1\",\"\") & ROW())-INDIRECT(SUBSTITUTE(ADDRESS(1,MATCH(\"Megérkezés időpontja\",$1:$1,0),4),\"1\",\"\") & ROW()))*24*60,0)"
          end
          
            write_item(worksheet, custom_data, item_index, column_index, cell_format, false, nil, hyperlink_format)
            width = get_column_width(custom_data)
            columns_width[column_index] = width if columns_width[column_index] < width
        end

        current_column_index = custom_columns.size

        original_columns.each_with_index do |c, column_index|
          value = xlsx_content(c, item)
          write_item(worksheet, value, item_index, current_column_index + column_index, cell_format, (c.name == :id), item.id, hyperlink_format)
          width = get_column_width(value)
          columns_width[column_index] = width if columns_width[column_index] < width
        end
      end
    end
    

    def xlsx_content(column, item)
      csv_content(column, item)
    end

    # Conditions from worksheet.rb in write_xlsx.
    def is_transformed_to_hyperlink?(token)
      return if not token.is_a?(String)
      # Match http, https or ftp URL
      if token =~ %r|\A[fh]tt?ps?://|
        true
        # Match mailto:
      elsif token =~ %r|\Amailto:|
        true
        # Match internal or external sheet link
      elsif token =~ %r!\A(?:in|ex)ternal:!
        true
      end
    end

    def crlf_to_lf(value)
      value.is_a?(String) ? value.gsub(/\r\n?/, "\n") : value
    end

    def write_item(worksheet, value, row_index, column_index, cell_format, is_id_column, id, hyperlink_format)
      if is_id_column
        issue_url = url_for(:controller => 'issues', :action => 'show', :id => id)
        worksheet.write(row_index + 1, column_index, issue_url, hyperlink_format, value)
        return
      end

      if is_transformed_to_hyperlink?(value)
        worksheet.write_string(row_index + 1, column_index, value, cell_format)
        return
      end

      worksheet.write(row_index + 1, column_index, crlf_to_lf(value), cell_format)
    end

    def get_column_width(value)
      value_str = value.to_s
      width = (value_str.length + value_str.chars.reject(&:ascii_only?).length) * 1.1  # 1.1: margin
      width > 30 ? 30 : width  # 30: max width
    end

    def create_header_format(workbook)
      workbook.add_format(:bold => 1,
                          :border => 1,
                          :color => 'white',
                          :bg_color => 'gray',
                          :text_wrap => 1,
                          :valign => 'top')
    end

    def create_cell_format(workbook)
      workbook.add_format(:border => 1,
                          :text_wrap => 1,
                          :valign => 'top')
    end

    def create_hyperlink_format(workbook)
      workbook.add_format(:border => 1,
                          :text_wrap => 1,
                          :valign => 'top',
                          :color => 'blue',
                          :underline => 1)
    end

  end
end