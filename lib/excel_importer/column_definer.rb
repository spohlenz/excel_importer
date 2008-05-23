class ExcelImporter::ColumnDefiner
  def initialize(importer)
    @importer = importer
  end
  
  def string(field)
    @importer.add_column(field, :string)
  end
  
  def text(field)
    @importer.add_column(field, :text)
  end
  
  def integer(field)
    @importer.add_column(field, :integer)
  end
  
  def date(field)
    @importer.add_column(field, :date)
  end
  
  def url(field)
    @importer.add_column(field, :url)
  end
  
  def custom(field, proc)
    @importer.add_column(field, proc)
  end
  
  def ignore(num_columns=1)
    num_columns.times { @importer.ignore_column }
  end
end
