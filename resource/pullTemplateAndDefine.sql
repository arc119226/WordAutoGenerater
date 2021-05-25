-- download config and template
select TEMPLATEFILENAME,
	   TemplateFileData,
	   CONFIGFILENAME,
	   ConfilgFileData 
from ReporterCategory 
 where REPORTCATEGORY='%s' 