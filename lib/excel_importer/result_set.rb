class ExcelImporter::ResultSet
  attr_accessor :all, :saved, :unsaved
  
  def initialize
    @all = []
    @saved = []
    @unsaved = []
  end
end
