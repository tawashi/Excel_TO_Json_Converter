require 'roo'
require 'fileutils'
require 'digest/md5'
require 'json'

rootDir = File.dirname(__FILE__) + "/../"

filename = ARGV.shift
unless filename
  STDERR.puts "Usage: ruby #{__FILE__} <filename> <env> <version>"
  exit
end

env = "develop"

baseDir = rootDir + env
jsonDir = rootDir + "json/" + env

FileUtils.mkdir_p(baseDir) unless FileTest.exist?(baseDir)
FileUtils.mkdir_p(jsonDir) unless FileTest.exist?(jsonDir)

filterFilePath = rootDir + "originData/" + env + "/json_list.json"

class Excel
  def initialize(filename)
    @excel = Roo::Excelx.new(filename)
  end

  def each_data_sheets
    @excel.sheets.select{ |s| s =~ /^[A-Za-z0-9_]+$/ }.each do |s|
      yield DataSheet.new(s, @excel.sheet(s))
    end
  end
end

class DataSheet
  TAG_SEARCH_ROW_COUNT = 100
  QUITTING_COLUMN_COUNT = 10
  QUITTING_ROW_COUNT = 50
  AVAILABLE_COLUMN_COUNT = 100
  AVAILABLE_ROW_COUNT = 100000

  def initialize(name, sheet)
    @name = name
    @sheet = sheet
  end

  def name
    @name
  end

  def rows
    empty_count = 0;
    rows = {}
    (data_start_row..AVAILABLE_ROW_COUNT).each do |i|
      if row = row(i)
        id = row['id']
        raise "ID can't be nil" if id.nil?
        raise "Duplicate id: #{id}" if rows[id]

        rows[id] = row
        empty_count = 0
      else
        break if (empty_count += 1) >= QUITTING_ROW_COUNT
      end
    end
    rows
  end

  private
  def row(row)
    data = {}
    columns.each do |col, meta|
      v = @sheet.cell(row, col)
      data[meta[:name]] = convert(v, meta[:type])
    end

    if data.all?{ |k, v| v.nil? || v == "" }
      return nil
    end

    data
  end

  def convert(value, type)
    if value.nil?
      if type == "string"
        return ""
      else 
        return nil
      end
    end

    case type
    when "intstring"
      return value.to_i.to_s
    when "string"
      return value.to_s if value.is_a?(String)
    when "int"
      return value.to_i if value.is_a?(Numeric)
    when "float"
      return value.to_f if value.is_a?(Float)

    when "bool"
      return value.to_i != 0 if value.is_a?(Numeric)
    when "datetime"
      return value.to_s if value =~ /^20\d{2}-[01]\d-[0-3]\d [0-2]\d:[0-5]\d:[0-5]\d$/
    else
      raise "Unsupported type: #{type}"
    end

    raise "Invalid value #{value} for #{type}"
  end

  def columns
    @columns ||= -> do
      empty_count = 0
      columns = {}
      (2..AVAILABLE_COLUMN_COUNT).each do |i|
        if meta = column_meta_data(i)
          columns[i] = meta
          empty_count = 0
        else
          break if (empty_count += 1) >= QUITTING_COLUMN_COUNT
        end
      end
      columns
    end.call
  end

  def column_meta_data(column)
    name = @sheet.cell(column_name_row, column)
    return nil unless name && name.length > 0

    type = @sheet.cell(data_type_row, column)
    raise "column #{name}: data type is not specified" unless type && type.length > 0

    {name: name, type: type}
  end

  def column_name_row
    @column_name_row ||= tag_row('column_name')
  end

  def data_type_row
    @data_type_row ||= tag_row('data_type')
  end

  def data_start_row
    @data_start_row ||= tag_row('data_start')
  end

  def tag_row(tag)
    (1..TAG_SEARCH_ROW_COUNT).each do |i|
      return i if @sheet.cell(i, 1) == tag
    end
    raise "tag #{tag} is not found in first #{TAG_SEARCH_ROW_COUNT} rows"
  end
end

class FilterList
  def initialize(filePath)
    if FileTest.exist?(filePath)
      @jsonArray = open(filePath) do |io|
        JSON.load(io)
      end
    end
  end

  def is_filter_check(name)
    if @jsonArray.include?(name)
      return true
    else
      return false
    end
  end

end

filterList = FilterList.new(filterFilePath)

Excel.new(filename).each_data_sheets do |sheet|
  begin
    rows = sheet.rows

      json = jsonDir + "/" + sheet.name + ".json"

      File.open(json, 'w') do |f2|
        f2.write(`php scripts/php_json_encode.php #{php}`)
      end

      hash = Digest::MD5.file(json).hexdigest.to_s
      size = File.size?(json)
      puts "#{sheet.name} hash:#{hash}"

    puts "#{sheet.name}: OK"

  rescue => e
    STDERR.puts "#{sheet.name}: #{e.message}"
  end

end

