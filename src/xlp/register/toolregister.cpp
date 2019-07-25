/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/register/register.hpp>


namespace xlp {

	void registerToolFunctions() {

		// xlpTime
		registerXLFunction(
			"XLPython - tools", 
			"xlpTime",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpTime",
			"trigger", 
			"get ellapsed time",
			"trigger");

		// xlpArgs
		registerXLFunction(
			"XLPython - tools", 
			"xlpArgs",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpArgs",
			"args01, args02, args03, args04, args05, transpose, trigger", 
			"create a vector of arguments", 
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"transpose\n"
			"trigger");

		// xlpBigArgs
		registerXLFunction(
			"XLPython - tools", 
			"xlpBigArgs",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpBigArgs",
			"args01, args02, args03, args04, args05, "
			"args06, args07, args08, args09, args10, "
			"args11, args12, args13, args14, args15, "
			"args16, args17, args18, transpose, trigger", 
			"create a vector of arguments", 
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"arguments\n"
			"transpose\n"
			"trigger");

		// xlpTrigger
		registerXLFunction(
			"XLPython - tools", 
			"xlpTrigger",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpTrigger",
			"trigger01, trigger02, trigger03, trigger04, trigger05", 
			"force dependencies between cells", 
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger");

		// xlpBigTrigger
		registerXLFunction(
			"XLPython - tools", 
			"xlpBigTrigger",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpBigTrigger",
			"trigger01, trigger02, trigger03, trigger04, trigger05, trigger06, trigger07, trigger08, trigger09, trigger10, "
			"trigger11, trigger12, trigger13, trigger14, trigger15, trigger16, trigger17, trigger18, trigger19, trigger20", 
			"force dependencies between cells", 
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger\n"
			"trigger");

		// xlpVolatile
		registerXLFunction(
			"XLPython - tools", 
			"xlpVolatile",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED+XL_VOLATILE,
			"xlpVolatile",
			"trigger", 
			"fix a cell as volatile", 
			"trigger");
 
		// xlpTranspose
		registerXLFunction(
			"XLPython - tools", 
			"xlpTranspose",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpTranspose",
			"scalars, trigger", 
			"transpose a scalars matrix", 
			"scalars\n"
			"trigger");

		// xlpT
		registerXLFunction(
			"XLPython - tools", 
			"xlpTranspose",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpT",
			"scalars, trigger", 
			"transpose a scalars matrix", 
			"scalars\n"
			"trigger");

		// xlpAdjustShape
		registerXLFunction(
			"XLPython - tools", 
			"xlpAdjustShape",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpAdjustShape",
			"scalars, rows, columns, transpose, trigger", 
			"adjust the shape of a matrix", 
			"scalars\n"
			"rows\n"
			"columns\n"
			"transpose ouput\n"
			"trigger");

		// xlpLoadFile
		registerXLFunction(
			"XLPython - tools", 
			"xlpLoadFile",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpLoadFile",
			"filename, id, trigger", 
			"load a script from file", 
			"filename\n"
			"id\n"
			"trigger");
	
		// xlpReduce
		registerXLFunction(
			"XLPython - tools", 
			"xlpReduce",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpReduce",
			"scalars, transpose, trigger", 
			"reduce a matrix of arguments (no error, no missing in full lines or columns)", 
			"scalars\n"
			"transpose output\n"
			"trigger");

		// xlpR
		registerXLFunction(
			"XLPython - tools", 
			"xlpReduce",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpR",
			"scalars, transpose, trigger", 
			"reduce a matrix of arguments (no error, no missing in full lines or columns)", 
			"scalars\n"
			"transpose output\n"
			"trigger");

		// xlpPretty
		registerXLFunction(
			"XLPython - tools", 
			"xlpPretty",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpPretty",
			"scalars, transpose, trigger", 
			"show a cell matrix in a pretty way", 
			"scalars\n"
			"transpose ooutput\n"
			"trigger");

		// xlpP
		registerXLFunction(
			"XLPython - tools", 
			"xlpPretty",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpP",
			"scalars, transpose, trigger", 
			"show a cell matrix in a pretty way", 
			"scalars\n"
			"transpose ooutput\n"
			"trigger");

		// xlpName
		/*registerXLFunction(
			"XLPython - tools", 
			"xlpName",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpName",
			"scalars ref, name, trigger", 
			"define a worksheet name", 
			"scalars reference\n"
			"name\n"
			"trigger");*/

		registerXLFunction(
			"XLPython - tools", 
			"xlpName",
			XL_LPOPER+XL_LPXLOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpName",
			"scalars ref, name, trigger", 
			"define a worksheet name", 
			"scalars reference\n"
			"name\n"
			"trigger");

		// xlpExcelId
		registerXLFunction(
			"XLPython - tools", 
			"xlpExcelId",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpExcelId",
			"sclars, trigger", 
			"return the excel id if possible", 
			"sclars\n trigger");

		// xlpPythonId
		registerXLFunction(
			"XLPython - tools", 
			"xlpPythonId",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpPythonId",
			"sclars, trigger", 
			"return the python id if possible", 
			"sclars\n trigger");

		// xlpSaveRange
		registerXLFunction(
			"XLPython - tools", 
			"xlpSaveRange",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpSaveRange",
			"scalars, id, trigger", 
			"save a range into user file", 
			"sclars\n id\n trigger");

		// xlpLoadRange
		registerXLFunction(
			"XLPython - tools", 
			"xlpLoadRange",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpLoadRange",
			"id, output, transpose, trigger", 
			"load a range into user file", 
			"id\n output\n transpose\n trigger");

		// xlpTmpDir
		registerXLFunction(
			"XLPython - tools", 
			"xlpTmpDir",
			XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpTmpDir",
			"trigger", 
			"get the temporary directory", 
			"trigger");

		// xlpFalseTrigger
		registerXLFunction(
			"XLPython - tools", 
			"xlpFalseTrigger",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpFalseTrigger",
			"scalar, false scalar, trigger", 
			"return false if scalar equal false scalar", 
			"scalar\nfalse scalar\ntrigger");

		// xlpoldValue
		registerXLFunction(
			"XLPython - tools", 
			"xlpSerie",
			XL_LPOPER+XL_LPXLOPER+XL_LPOPER+XL_UNCALCED,
			"xlpSerie",
			"scalar reference, trigger", 
			"return the serie values of a reference", 
			"scalar reference\ntrigger");

		// xlpSerieValue
		registerXLFunction(
			"XLPython - tools", 
			"xlpSerieValue",
			XL_LPOPER+XL_LPXLOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpSerieValue",
			"scalar reference, index,  trigger", 
			"return the serie value of a reference at index", 
			"scalar reference\nindex\ntrigger");

		// xlpSerieValue
		registerXLFunction(
			"XLPython - tools", 
			"xlpVisible",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPXLOPER+XL_LPOPER+XL_UNCALCED,
			"xlpVisible",
			"sheet, sheets, flag,  trigger", 
			"set worksheet visible or not", 
			"sheet\nsheets\nflag\ntrigger");

	}
	
}