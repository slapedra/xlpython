/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/register/register.hpp>


namespace xlp {

	void registerIOFunctions() {

		// xlpFullIO
		registerXLFunction(
			"XLPython - IO", 
			"xlpFullIO",
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED,
			"xlpFullIO",
			"tranpose, trigger", 
			"return full io output", 
			"transpose output data\n"
			"trigger");

		// xlpIO
		registerXLFunction(
			"XLPython - IO", 
			"xlpIO", 
			XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_LPOPER+XL_UNCALCED+XL_VOLATILE,
			"xlpIO",
			"nb lines, tranpose, trigger", 
			"return io output", 
			"number of lines\n"
			"transpose output data\n"
			"trigger");
		
	}
	
}