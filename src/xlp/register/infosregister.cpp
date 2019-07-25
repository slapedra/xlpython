/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/register/register.hpp>

namespace xlp {

	void registerInfosFunctions() {

		// xlpVersion
		registerXLFunction(
			"XLPython - infos", 
			"xlpVersion",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpVersion",
			"trigger", 
			"returns the version number of xlp", 
			"trigger");

		// xlpAuthor
		registerXLFunction(
			"XLPython - infos", 
			"xlpAuthor",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpAuthor",
			"trigger", 
			"returns the author name", 
			"trigger");

		// xlpContributor
		registerXLFunction(
			"XLPython - infos", 
			"xlpContributors",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpContributors",
			"trigger", 
			"returns the contributors name", 
			"trigger");

		// xlpPythonPath
		registerXLFunction(
			"XLPython - infos",
			"xlpPythonPath",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpPythonPath",
			"trigger", 
			"environnment variable PYTHONPATH", 
			"trigger");

		// xlpSysPath
		registerXLFunction(
			"XLPython - infos",
			"xlpSysPath",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpSysPath",
			"transpose, trigger", 
			"get the system paths", 
			"transpose\n"
			"trigger");


		// xlpPythonVersion
		registerXLFunction(
			"XLPython - infos", 
			"xlpPythonVersion",	
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpPythonVersion",
			"trigger", 
			"python version", 
			"trigger");

		// xlpWorkbookPath
		registerXLFunction(
			"XLPython - infos", 
			"xlpWorkbookPath",	
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpWorkbookPath",
			"trigger", 
			"Workbook path", 
			"trigger");

		// xlpWorkbookName
		registerXLFunction(
			"XLPython - infos", 
			"xlpWorkbookName",	
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpWorkbookName",
			"trigger", 
			"Workbook Name", 
			"trigger");

		// xlpWoksheetName
		registerXLFunction(
			"XLPython - infos", 
			"xlpWorksheetName",	
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpWorksheetName",
			"little, trigger", 
			"Worksheet name",
			"little format\n"
			"trigger");

		// xlpWoksheetProtected
		registerXLFunction(
			"XLPython - infos", 
			"xlpWorksheetProtected",	
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpWorksheetProtected",
			"trigger", 
			"is protected",
			"trigger");

	}
	
}