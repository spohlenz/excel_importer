class ExcelImporter::ResultSet
  attr_accessor :all, :saved, :unsaved
  
  def initialize
    @all = []
    @saved = []
    @unsaved = []
  end
  
  def save(instance)
    if instance.save
      saved << instance
    else
      unsaved << instance
    end
    
    all << instance
  rescue Mysql::Error => e
    raise ExcelImporter::ImportException
  end
end
