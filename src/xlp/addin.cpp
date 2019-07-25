/*
 2008 Sebastien Lapedra
*/
#include <xlp/config.hpp>
#include <xlp/xlpdefines.hpp>
#include <xlp/xl/xl.hpp>
#include <sstream>
#include <xlp/python/xlpsession.hpp>
#include <xlp/register/register.hpp>
#include <xlp/version.hpp>
#include <boost/bind.hpp>

using namespace xlp;

static bool initialized = false;
static bool XLP_remove = false;

// boost::bind does not work
void onWorkbookClose(const std::string& name) {
	XLPSession::instance().onWorkbookClose(name);
}

XL_EXPORT int xlAutoOpen() {		
	try {
		if (!initialized) {
			XLSession::instance().init();
			XLPSession::instance().init();
			registerIOFunctions();
			registerToolFunctions();
			registerObjectManagementFunctions();
			registerInfosFunctions();
			registerPythonExecutionFunctions();
			XLSession::instance().registerSession();
			XLSession::instance().setOnWorkbookClose(&onWorkbookClose);
			XLSession::instance().setNameKey(XLP_NAMEKEY);
			FreeAllTempMemory();
			initialized = true;
		}
		return 1;
	} catch (const Error &e) {
		std::ostringstream err;
		err << "Error loading Addin: " << e.what();
		Excel(xlcAlert, 0, 1, TempStrStl(err.str()));
		return 0;
	} catch (...) {
		return 0;
	}	
}


XL_EXPORT int xlAutoClose() {
	try {
		if ( (initialized) && (XLP_remove) ) {
			XLSession::instance().unregisterSession();
			initialized = false;
			XLP_remove = false;
		}	
		return 1;
	} catch (const Error &e) {
			std::ostringstream err;
			err << "Error unloading XLPython Addin: " << e.what();
	    Excel(xlcAlert, 0, 1, TempStrStl(err.str()));
    	return 0;
	} catch (...) {
			return 0;
	}
}

XL_EXPORT int xlAutoRemove() {
	try {
		XLP_remove = true;
		return 1;
	} catch (const Error &e) {
			std::ostringstream err;
			err << "Error unloading XLPython Addin: " << e.what();
	    Excel(xlcAlert, 0, 1, TempStrStl(err.str()));
    	return 0;
	} catch (...) {
			return 0;
	}
}

XL_EXPORT void xlAutoFree(XLOPER *px) {
	freeOper(px);
}

XL_EXPORT XLOPER *xlAddInManagerInfo(XLOPER *xlAction) {
	xlAutoOpen();
	XLOPER xlReturn;
	Excel(xlCoerce, &xlReturn, 2, xlAction, TempInt(xltypeInt));
	if (1 == xlReturn.val.w) {
		Excel(xlFree, 0, 1, &xlReturn);
		XlpOper ret(XLP_PRODUCTNAME);
		return ret.get();
  } else {
		Excel(xlFree, 0, 1, &xlReturn);
		return TempErr(xlerrNA);
	}
}


