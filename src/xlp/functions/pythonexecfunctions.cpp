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


namespace xlp {

	XL_EXPORT OPER *xlpSet(OPER* xlScalar, OPER* xlArgs, OPER* xlValues, OPER* xlIndex, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", true, false);
		XlpOper xlpId(xlId, "id", true, false);
		XlpOper xlpArgs(xlArgs, "args");
		XlpOper xlpValues(xlValues, "values", false, false, true);
		XlpOper xlpIndex(xlIndex, "index", false, false, true);
		Scalar scalar = xlpScalar;
		std::string id = xlpId;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		boost::shared_ptr< Range<Scalar> > values = xlpValues;
		boost::shared_ptr< Range<Scalar> > index = xlpIndex;
		std::string obj = XLPSession::instance().set(scalar, args, values, index, id);
		XL_END(obj)
	}

	XL_EXPORT OPER *xlpGet(OPER* xlScalar, OPER* xlArgs, OPER* xlOutput, OPER* xlForce, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_OUTPUT(xlTrigger, xlTranspose, xlOutput) 
		bool force = XLSession::instance().convertBool(xlForce, false);
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpArgs(xlArgs, "args");
		Scalar scalar = xlpScalar;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		boost::shared_ptr< Range<Scalar> > datas = XLPSession::instance().get(scalar, args, force);
		XL_END_OUTPUT(datas)
	}

	XL_EXPORT OPER *xlpExec(OPER* xlCode, OPER* xlArgs, OPER* xlIds, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpCode(xlCode, "code");
		XlpOper xlpIds(xlIds, "ids");
		XlpOper xlpArgs(xlArgs, "args");
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		boost::shared_ptr< Range<Scalar> > code = xlpCode;
		boost::shared_ptr< Range<Scalar> > ids = xlpIds;
		boost::shared_ptr< Range<Scalar> > objs = XLPSession::instance().exec(code, args, ids);
		boost::shared_ptr< Range<Scalar> > datas;
		if (objs->size() == 0) {
			datas.reset(new Range<Scalar>(1, 1));
			datas->data[0] = "success";
		} else {
			datas = objs;
		}
		XL_END(datas)
	}

	XL_EXPORT OPER *xlpCompile(OPER* xlCode, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpCode(xlCode, "code");
		XlpOper xlpId(xlId, "ids");
		std::string id = xlpId;
		boost::shared_ptr< Range<Scalar> > code = xlpCode;
		XLPSession::instance().compile(code, id);
		XL_END("success")
	}

	XL_EXPORT OPER *xlpImport(OPER* xlScalar, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", true, false);
		std::string module = xlpScalar;
		XlpOper xlpId(xlId, "id", true, false);
		std::string id = xlpId;
		std::string key = XLPSession::instance().import(module, id);
		XL_END(key)
	}

	XL_EXPORT OPER *xlpFromImport(OPER* xlScalar, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", true, false);
		std::string module = xlpScalar;
		XLPSession::instance().fromImport(module);
		XL_END("success")
	}

	XL_EXPORT OPER *xlpAttr(OPER* xlScalar, OPER* xlAttr, OPER* xlArgs, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpId(xlId, "id", false, false);
		XlpOper xlpAttr(xlAttr, "attr", false, false);
		XlpOper xlpArgs(xlArgs, "args");
		Scalar scalar = xlpScalar;
		std::string id = xlpId;
		std::string attr = xlpAttr;
		std::string key;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		key = XLPSession::instance().attr(scalar, attr, args, id);
		XL_END(key)
	}

	XL_EXPORT OPER *xlpGetAttr(OPER* xlScalar, OPER* xlAttr, OPER* xlArgs, 
		OPER* xlOutput,  OPER* xlForce, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_OUTPUT(xlTrigger, xlTranspose, xlOutput) 
		bool force = XLSession::instance().convertBool(xlForce, false);
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpAttr(xlAttr, "attr", false, false);
		XlpOper xlpArgs(xlArgs, "args");
		Scalar scalar = xlpScalar;
		std::string attr = xlpAttr;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		boost::shared_ptr< Range<Scalar> > datas = XLPSession::instance().getAttr(scalar, attr, args, force);
		XL_END_OUTPUT(datas)
	}

	XL_EXPORT OPER *xlpApply(OPER* xlScalar, OPER* xlArgs, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpId(xlId, "id", false, false);
		XlpOper xlpArgs(xlArgs, "args");
		Scalar scalar = xlpScalar;
		std::string id = xlpId;
		std::string key;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		key = XLPSession::instance().apply(scalar, args, id);
		XL_END(key)
	}

	XL_EXPORT OPER *xlpGetApply(OPER* xlScalar, OPER* xlArgs, 
		OPER* xlOutput,  OPER* xlForce, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_OUTPUT(xlTrigger, xlTranspose, xlOutput) 
		bool force = XLSession::instance().convertBool(xlForce, false);
		XlpOper xlpScalar(xlScalar, "scalar", false, false);
		XlpOper xlpArgs(xlArgs, "args");
		Scalar scalar = xlpScalar;
		boost::shared_ptr< Range<Scalar> > args = xlpArgs;
		boost::shared_ptr< Range<Scalar> > datas = XLPSession::instance().getApply(scalar, args, force);
		XL_END_OUTPUT(datas);
	}
	
	XL_EXPORT OPER *xlpInsertPaths(OPER* xlPaths, OPER * xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpPaths(xlPaths, "scalar", false, false);
		boost::shared_ptr< Range<Scalar> > scalars = xlpPaths;
		std::vector<std::string> paths;
		for (Size i=0;i<scalars->data.size();i++) {
			paths.push_back(scalars->data[i]);
		}
		XLPSession::instance().insertPaths(paths);
		XL_END("success")
	}

}



