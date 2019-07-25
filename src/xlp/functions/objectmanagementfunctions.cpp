/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <limits>
#include <algorithm>
#include <xlp/xlpdefines.hpp>
#include <xlp/python/xlpsession.hpp>
#include <boost/timer.hpp>
#include <xlp/xl/xl.hpp>


namespace xlp {

	XL_EXPORT OPER *xlpType(OPER* xlScalar, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		Scalar scalar = xlpScalar;
		XL_END(XLPSession::instance().type(scalar))
	}

	XL_EXPORT OPER *xlpListAllObjects(OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		boost::shared_ptr< Range<Scalar> > objs = XLPSession::instance().listAllObjects();
		XL_END(objs)
	}

	XL_EXPORT OPER *xlpDeleteAllObjects(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XLPSession::instance().deleteAllObjects();
		XL_END("success")
	}

	XL_EXPORT OPER *xlpDeleteObject(OPER* xlScalar, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		Scalar scalar =xlpScalar;
		XLPSession::instance().deleteObject(scalar);
		XL_END("success")
	}
	

	XL_EXPORT OPER *xlpListAllAttr(OPER* xlScalar, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		Scalar scalar = xlpScalar;
		boost::shared_ptr< Range<Scalar> > attrs = XLPSession::instance().listAllAttr(scalar);
		XL_END(attrs)
	}

	XL_EXPORT OPER *xlpAttrInfos(OPER* xlScalar, OPER* xlAttr, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpAttr(xlAttr, "attr", false, false);
		Scalar scalar = xlpScalar;
		std::string attr = xlpAttr;
		boost::shared_ptr< Range<Scalar> > infos = XLPSession::instance().attrInfos(scalar, attr);
		XL_END(infos)
	}

	XL_EXPORT OPER *xlpWorkbook(OPER* xlScalar, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		std::string id = xlpScalar;
		XL_END(XLPSession::instance().workbook(id))
	}

}



