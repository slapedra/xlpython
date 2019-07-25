/*
 2008 Sebastien Lapedra
*/

#ifndef xldefines_hpp
#define xldefines_hpp

#define NOMINMAX
#include <windows.h>
#include <xlp/xl/xlcall.h>


namespace xlp {

	#define DLL_API			__declspec(dllexport)
	#define XL_EXPORT		extern 	"C" __declspec(dllexport)

	#define XL_MAX_STR_LEN		256

	#define XL_TIMESTEP			1

	// parameters registered with Excel as OPER (P) are declared as XLOPER
	#define OPER XLOPER

	typedef struct {
	    WORD rows;
	    WORD columns;
	    double array[1];
	} FP;

	// excel arguments type
	// generally only "P" is used
	#define XL_BOOL					std::string("A")	// short int used as logical
	#define XL_DOUBLE				std::string("B")	// double
	#define XL_CSTRING			std::string("C")	// char * to C style NULL terminated string (up to 255 characters)
	#define XL_PSTRING			std::string("D")	// unsigned char * to Pascal style byte counted string (up to 255 characters)
	#define XL_DOUBLE_			std::string("E") 	// pointer to double
	#define XP_CSTRING_			std::string("F") 	// reference to a null terminated string
	#define XL_PSTRING_			std::string("G")	// reference to a byte counted string
	#define XL_USHORT				std::string("H")	// unsigned 2 byte int
	#define XL_SHORT				std::string("I") 	// signed 2 byte int
	#define XL_INT      		std::string("I") 	// signed 2 byte int - platform dependent
	#define XL_LONG					std::string("J")	// signed 4 byte int
	#define XL_ARRAY				std::string("K")	// pointer to struct FP
	#define XL_BOOL_				std::string("L")	// reference to a boolean
	#define XL_SHORT_				std::string("M") 	// reference to signed 2 byte int
	#define XL_LONG_				std::string("N")	// reference to signed 4 byte int
	#define XL_ARRAY_				std::string("O")	// three arguments are ushort*, ushort*, double[]
	#define XL_LPOPER   		std::string("P")	// pointer to OPER struct (never a reference type).
	#define XL_RANGE    		std::string("P")	// alias for LPOPER
	#define XL_LPXLOPER			std::string("R")	// pointer to XLOPER struct
	#define XL_LPSREF				std::string("R")	// pointer to SREF struct
	#define XL_LPMREF				std::string("R")	// pointer to MREF struct

	// add in function behaviour.
	#define XL_VOLATILE			std::string("!")	// called every time sheet is recalced
	#define XL_UNCALCED 		std::string("#")	// dereferencing uncalced cells returns old value


	#define XL_BEGIN(x)\
	try {\
		xlp::XLSession::instance().enterFunction(x);

	#define XL_BEGIN_TRANSPOSE(x, y)\
	try {\
		xlp::XLSession::instance().enterFunction(x, y);

	#define XL_BEGIN_OUTPUT(x, y, z)\
	try {\
		xlp::XLSession::instance().enterFunction(x, y, z);

	#define XL_END(x)\
		{\
			XlpOper ret(x, "", XLSession::instance().transpose());\
			xlp::XLSession::instance().exitFunction();\
			return ret.get();\
		}\
	} catch (Error &e) { \
		std::cerr << e.message() <<std::endl; \
		XlpOper ret(NullType(), "");\
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} catch (...) { \
		std::string msg("Unknown exception"); \
		std::cerr << msg << std::endl; \
		XlpOper ret(NullType(), "");\
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} 

	#define XL_END_OUTPUT(x)\
		{\
			std::string res = XLSession::instance().fillOutput(x);\
			if (res.empty()) {\
				XlpOper ret(x, "", XLSession::instance().transpose());\
				xlp::XLSession::instance().exitFunction();\
				return ret.get();\
			} else {\
				XlpOper ret(res, "");\
				xlp::XLSession::instance().exitFunction();\
				return ret.get();\
			}\
		}\
	} catch (Error &e) { \
		std::cerr << e.message() <<std::endl; \
		try {\
			boost::shared_ptr< Range<Scalar> > datas(new Range<Scalar>(1,1)); \
			datas->data[0] = NAType();\
			xlp::XLSession::instance().fillOutput(datas); \
		} catch(Error&) {\
		}\
		XlpOper ret(NullType(), ""); \
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} catch (...) { \
		std::string msg("Unknown exception"); \
		std::cerr << msg << std::endl; \
		XlpOper ret(NullType(), ""); \
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} 

	#define XL_END_PRETTY(x)\
		{\
			XlpOper ret(x, "", XLSession::instance().transpose(), true);\
			xlp::XLSession::instance().exitFunction();\
			return ret.get();\
		}\
	} catch (Error &e) { \
		std::cerr << e.message() <<std::endl; \
		XlpOper ret(NullType(), "");\
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} catch (...) { \
		std::string msg("Unknown exception"); \
		std::cerr << msg << std::endl; \
		XlpOper ret(NullType(), "");\
		xlp::XLSession::instance().exitFunction();\
		return ret.get(); \
	} 

}

#endif

