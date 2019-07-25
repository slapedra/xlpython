/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <limits>
#include <algorithm>
#include <xlp/xl/xl.hpp>
#include <xlp/xlpdefines.hpp>
#include <xlp/python/xlpsession.hpp>
#include <boost/timer.hpp>
#include <xlp/version.hpp>



namespace xlp {

	XL_EXPORT OPER *xlpVersion(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLP_STR_PRODUCTVERSION)
	}


	XL_EXPORT OPER *xlpAuthor(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END("Sebastien Lapedra")
	}

	XL_EXPORT OPER *xlpContributors(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END("None")
	}
	
	XL_EXPORT OPER *xlpPythonPath(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLPSession::instance().pythonPath())
	}

	XL_EXPORT OPER *xlpSysPath(OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XL_END(XLPSession::instance().sysPath())
	}

	XL_EXPORT OPER *xlpPythonVersion(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLPSession::instance().pythonVersion())
	}

	XL_EXPORT OPER *xlpWorkbookPath(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLPSession::instance().workbookPath())
	}

	XL_EXPORT OPER *xlpWorkbookName(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLPSession::instance().workbookName())
	}

	XL_EXPORT OPER *xlpWorksheetName(OPER* xlLittle, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		bool little = XLSession::instance().convertBool(xlLittle, true);
		XL_END(XLPSession::instance().worksheetName(little))
	}

	XL_EXPORT OPER *xlpWorksheetProtected(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		bool ret = XLSession::instance().isProtected();
		XL_END(ret)
	}

}



