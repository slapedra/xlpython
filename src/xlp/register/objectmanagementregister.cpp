/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/register/register.hpp>


namespace xlp {

	void registerObjectManagementFunctions() {

		// xlpType
		registerXLFunction(
			"XLPython - objects", 
			"xlpType",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpType",
			"scalar, trigger", 
			"return the type of a scalar or an object", 
			"scalar \n trigger");
		 
		// xlpListAllObjects
		registerXLFunction(
			"XLPython - objects", 
			"xlpListAllObjects",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpListAllObjects",
			"transpose, trigger", 
			"return a list of all python objects", 
			"transpose output data\n"
			"trigger");

		// xlpDeleteObject
		registerXLFunction(
			"XLPython - objects", 
			"xlpDeleteObject",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpDeleteObject",
			"scalar, trigger", 
			"return a python object", 
			"scalar\n"
			"trigger");

		// xlpDeleteAllObjects
		registerXLFunction(
			"XLPython - objects", 
			"xlpDeleteAllObjects",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpDeleteAllObjects",
			"trigger", 
			"delete all python objects", 
			"trigger");

		// xlpListAllAttr
		registerXLFunction(
			"XLPython - objects", 
			"xlpListAllAttr",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpListAllAttr",
			"scalar, transpose, trigger", 
			"list all the attributes of a scalar/object", 
			"scalar\n"
			"transpose output data\n"
			"trigger");

		// xlpAttrInfos
		registerXLFunction(
			"XLPython - objects", 
			"xlpAttrInfos",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpAttrInfos",
			"scalar, attribute, transpose, trigger", 
			"get attribute infos", 
			"scalar\n"
			"attribute name\n"
			"transpose\n"
			"trigger");

		// xlpWorkbook
		registerXLFunction(
			"XLPython - objects", 
			"xlpWorkbook",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpWorkbook",
			"id, trigger", 
			"return the workbook name of an object", 
			"ôbject id \n trigger");

	}
	
}