/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/register/register.hpp>


namespace xlp {

	void registerPythonExecutionFunctions() {

		// xlpSet
		registerXLFunction(
			"XLPython - core", 
			"xlpSet", 
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpSet",
			"scalar, args, values, index, id, trigger", 
			"set a python object", 
			"scalar or python code\n"
			"arguments\n"
			"values for inserting\n"
			"index for values (for dict type uses)\n"
			"id name of the object\n"
			"trigger");
		
		// xlpGet
		registerXLFunction(
			"XLPython - core", 
			"xlpGet", 
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpGet",
			"scalar, args, output, force, transpose, trigger", 
			"get a python object", 
			"scalar or python code\n"
			"arguments\n"	
			"output cell matrix id\n"
			"force get\n"
			"transpose output data\n"
			"trigger");
 
		// xlpExec
		registerXLFunction(
			"XLPython - core", 
			"xlpExec",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpExec",
			"code, args, ids, transpose, trigger", 
			"execute python code", 
			"python code\n"
			"arguments\n"
			"ids object\n"
			"transpose\n"
			"trigger");

		// xlpExec
		registerXLFunction(
			"XLPython - core", 
			"xlpCompile",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpCompile",
			"code, id, trigger", 
			"compile python code", 
			"python code\n"
			"id\n"
			"trigger");

		// xlpImport
		registerXLFunction(
			"XLPython - core", 
			"xlpImport",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpImport",
			"module, id, trigger", 
			"import a module as an object", 
			"module\n"
			"user id of the object\n"
			"trigger");

		// xlpFromImport
		registerXLFunction(
			"XLPython - core", 
			"xlpFromImport",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpFromImport",
			"module, trigger", 
			"from module import *", 
			"module\n"
			"trigger");

	
		// xlpAttr
		registerXLFunction(
			"XLPython - core", 
			"xlpAttr",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpAttr",
			"scalar, attribute, args, id, trigger", 
			"store an atttribute as object", 
			"scalar\n"
			"attribute name\n"
			"arguments\n"
			"id\n"
			"trigger");

		// xlpGetAttr
		registerXLFunction(
			"XLPython - core", 
			"xlpGetAttr",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpGetAttr",
			"scalar, attribute, args, output, force, transpose, trigger", 
			"get an atttribute", 
			"scalar\n"
			"attribute name\n"
			"arguments\n"
			"output cell matrix\n"
			"force get\n"
			"transpose ouput datas\n"
			"trigger");
	
		// xlpApply
		registerXLFunction(
			"XLPython - core", 
			"xlpApply",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpApply",
			"scalar, args, id, trigger", 
			"apply a function", 
			"scalar\n"
			"arguments\n"
			"id\n"
			"trigger");

		// xlpGetApply
		registerXLFunction(
			"XLPython - core", 
			"xlpGetApply",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpGetApply",
			"scalar, args, output, force, transpose, trigger", 
			"get the result of an apply", 
			"scalar\n"
			"arguments\n"
			"output cell matrix id\n"
			"force get\n"
			"transpose ouput data\n"
			"trigger");

		// xlpInsertPaths
		registerXLFunction(
			"XLPython - core", 
			"xlpInsertPaths",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpInsertPaths",
			"paths, trigger", 
			"insert a list of path into the system path", 
			"paths\n"
			"trigger");

	}
	
}