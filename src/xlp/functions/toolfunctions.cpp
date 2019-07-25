/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <limits>
#include <iostream>
#include <algorithm>
#include <xlp/xl/xl.hpp>
#include <xlp/xlpdefines.hpp>
#include <xlp/python/xlpsession.hpp>
#include <boost/timer.hpp>



namespace xlp {

	XL_EXPORT OPER *xlpTime(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLSession::instance().time())
	}

	XL_EXPORT OPER *xlpArgs(
		OPER* xlArgs01, OPER* xlArgs02, OPER* xlArgs03, OPER* xlArgs04, OPER* xlArgs05, 
		OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		std::vector<XlpOper> xlpArgs;
		xlpArgs.push_back(XlpOper(xlArgs01, "args01"));
		xlpArgs.push_back(XlpOper(xlArgs02, "args02"));
		xlpArgs.push_back(XlpOper(xlArgs03, "args03"));
		xlpArgs.push_back(XlpOper(xlArgs04, "args04"));
		xlpArgs.push_back(XlpOper(xlArgs05, "args05"));
		std::vector< boost::shared_ptr< Range<Scalar> > > args;
		Size size = 0;
		for (Size i=0;i<xlpArgs.size();i++) {
			boost::shared_ptr< Range<Scalar> > tmp = xlpArgs[i];
			args.push_back(tmp);
			size += args[i]->size();
		}
		boost::shared_ptr< Range<Scalar> > fullArgs(new Range<Scalar>(1, size));
		Size offset = 0;
		for (Size i=0;i<args.size();i++) {
			for (Size j=0;j<args[i]->size();j++) {
				fullArgs->data[offset] = args[i]->data[j];
				offset++;
			}
		}
		XL_END(fullArgs)
	}

	XL_EXPORT OPER *xlpBigArgs(
		OPER* xlArgs01, OPER* xlArgs02, OPER* xlArgs03, OPER* xlArgs04, OPER* xlArgs05,
		OPER* xlArgs06, OPER* xlArgs07, OPER* xlArgs08, OPER* xlArgs09, OPER* xlArgs10, 
		OPER* xlArgs11, OPER* xlArgs12, OPER* xlArgs13, OPER* xlArgs14, OPER* xlArgs15, 
		OPER* xlArgs16, OPER* xlArgs17, OPER* xlArgs18, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		std::vector<XlpOper> xlpArgs;
		xlpArgs.push_back(XlpOper(xlArgs01, "args01"));
		xlpArgs.push_back(XlpOper(xlArgs02, "args02"));
		xlpArgs.push_back(XlpOper(xlArgs03, "args03"));
		xlpArgs.push_back(XlpOper(xlArgs04, "args04"));
		xlpArgs.push_back(XlpOper(xlArgs05, "args05"));
		xlpArgs.push_back(XlpOper(xlArgs06, "args06"));
		xlpArgs.push_back(XlpOper(xlArgs07, "args07"));
		xlpArgs.push_back(XlpOper(xlArgs08, "args08"));
		xlpArgs.push_back(XlpOper(xlArgs09, "args09"));
		xlpArgs.push_back(XlpOper(xlArgs10, "args10"));
		xlpArgs.push_back(XlpOper(xlArgs11, "args11"));
		xlpArgs.push_back(XlpOper(xlArgs12, "args12"));
		xlpArgs.push_back(XlpOper(xlArgs13, "args13"));
		xlpArgs.push_back(XlpOper(xlArgs14, "args14"));
		xlpArgs.push_back(XlpOper(xlArgs15, "args15"));
		xlpArgs.push_back(XlpOper(xlArgs16, "args16"));
		xlpArgs.push_back(XlpOper(xlArgs17, "args17"));
		xlpArgs.push_back(XlpOper(xlArgs18, "args18"));
		std::vector< boost::shared_ptr< Range<Scalar> > > args;
		Size size = 0;
		for (Size i=0;i<xlpArgs.size();i++) {
			boost::shared_ptr< Range<Scalar> > tmp = xlpArgs[i];
			args.push_back(tmp);
			size += args[i]->size();
		}
		boost::shared_ptr< Range<Scalar> > fullArgs(new Range<Scalar>(1, size));
		Size offset = 0;
		for (Size i=0;i<args.size();i++) {
			for (Size j=0;j<args[i]->size();j++) {
				fullArgs->data[offset] = args[i]->data[j];
				offset++;
			}
		}
		XL_END(fullArgs)
	}

	XL_EXPORT OPER *xlpTrigger(
		OPER* xlTrig01, OPER* xlTrig02, OPER* xlTrig03, OPER* xlTrig04, OPER* xlTrig05) {
		try {
			boost::shared_ptr< Range<Scalar> > datas(new Range<Scalar>(1, 1));
			static Integer index = 0;
			try {
				xlp::XLSession::instance().validateTrigger(xlTrig01);
				xlp::XLSession::instance().validateTrigger(xlTrig02);
				xlp::XLSession::instance().validateTrigger(xlTrig03);
				xlp::XLSession::instance().validateTrigger(xlTrig04);
				xlp::XLSession::instance().validateTrigger(xlTrig05);
				++index;
				datas->data[0] = index;
			} catch (Error&) {
				datas->data[0] = false;	
			}
		XL_END(datas)
	}

	XL_EXPORT OPER *xlpBigTrigger(
		OPER* xlTrig01, OPER* xlTrig02, OPER* xlTrig03, OPER* xlTrig04, OPER* xlTrig05, 
		OPER* xlTrig06, OPER* xlTrig07, OPER* xlTrig08, OPER* xlTrig09, OPER* xlTrig10,
		OPER* xlTrig11, OPER* xlTrig12, OPER* xlTrig13, OPER* xlTrig14, OPER* xlTrig15, 
		OPER* xlTrig16, OPER* xlTrig17, OPER* xlTrig18, OPER* xlTrig19, OPER* xlTrig20) {
		try {
			boost::shared_ptr< Range<Scalar> > datas(new Range<Scalar>(1, 1));
			static Integer index = 0;
			try {
				xlp::XLSession::instance().validateTrigger(xlTrig01);
				xlp::XLSession::instance().validateTrigger(xlTrig02);
				xlp::XLSession::instance().validateTrigger(xlTrig03);
				xlp::XLSession::instance().validateTrigger(xlTrig04);
				xlp::XLSession::instance().validateTrigger(xlTrig05);
				xlp::XLSession::instance().validateTrigger(xlTrig06);
				xlp::XLSession::instance().validateTrigger(xlTrig07);
				xlp::XLSession::instance().validateTrigger(xlTrig08);
				xlp::XLSession::instance().validateTrigger(xlTrig09);
				xlp::XLSession::instance().validateTrigger(xlTrig10);
				xlp::XLSession::instance().validateTrigger(xlTrig11);
				xlp::XLSession::instance().validateTrigger(xlTrig12);
				xlp::XLSession::instance().validateTrigger(xlTrig13);
				xlp::XLSession::instance().validateTrigger(xlTrig14);
				xlp::XLSession::instance().validateTrigger(xlTrig15);
				xlp::XLSession::instance().validateTrigger(xlTrig16);
				xlp::XLSession::instance().validateTrigger(xlTrig17);
				xlp::XLSession::instance().validateTrigger(xlTrig18);
				xlp::XLSession::instance().validateTrigger(xlTrig19);
				xlp::XLSession::instance().validateTrigger(xlTrig20);
				++index;
				datas->data[0] = index;
			} catch (Error&) {
				datas->data[0] = false;
			}
		XL_END(datas)
	}

	XL_EXPORT OPER *xlpTranspose(OPER* xlScalars, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpTrigger(xlTrigger, "trigger");
		XlpOper xlpScalars(xlScalars, "scalars", true);
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		XL_END(scalars)
	}

	XL_EXPORT OPER *xlpAdjustShape(OPER* xlScalars, OPER* xlRows, OPER* xlColumns, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpColumns(xlColumns, "columns");
		XlpOper xlpRows(xlRows, "rows");
		XlpOper xlpScalars(xlScalars, "scalars");
		Integer rows = xlpRows;
		Integer columns = xlpColumns;	
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		scalars->rows = rows;
		scalars->columns = columns;
		REQUIRE (Size(columns*rows) == scalars->size(), "bad size for new shape: " 
			<< rows*columns << " != " << scalars->size());
		XL_END(scalars)
	}

	XL_EXPORT OPER *xlpLoadFile(OPER* xlFile, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpFile(xlFile, "filename", false, false);
		XlpOper xlpId(xlId, "id", true, false);
		std::string filename = xlpFile;
		std::string id = xlpId;
		std::string key = XLPSession::instance().loadFile(filename, id);
		XL_END(key)
	}

	XL_EXPORT OPER *xlpVolatile(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		static Integer index = 0;
		index++;
		XL_END(index)
	}

	XL_EXPORT OPER *xlpReduce(OPER* xlArgs, OPER* xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpArgs(xlArgs, "args", false, false, true);
		boost::shared_ptr< Range<Scalar> > res = xlpArgs;
		XL_END(res)
	}

	XL_EXPORT OPER *xlpPretty(OPER* xlScalars, OPER* xlTranspose, OPER * xlTrigger) {
		XL_BEGIN_TRANSPOSE(xlTrigger, xlTranspose)
		XlpOper xlpScalars(xlScalars, "scalars");
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		XL_END_PRETTY(scalars)
	}

	XL_EXPORT OPER *xlpName(OPER* xlRef, OPER* xlName, OPER * xlTrigger) {
		XL_BEGIN(xlTrigger)
		static XLOPER xlBool;
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = false;
		Excel4(xlfVolatile, 0, 1, &xlBool);
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		boost::shared_ptr<Xloper> xllRef(new Xloper);
		int result = Excel4(xlCoerce, &(*xllRef), 2, xlRef, &xlInt);
		Scalar ret;
		if (result != xlretSuccess) {
			ret = NullType();
			std::cerr << "filling warning message error" << std::endl;
		} else {
			ret = "success";
			boost::shared_ptr<XLRef> ref(new XLRef(xllRef));
			XlpOper xlpName(xlName, "name");
			std::string name = xlpName;
			XLSession::instance().defineName(ref, name);
		}
		XL_END(ret)
	}

	XL_EXPORT OPER *xlpExcelId(OPER* xlScalars, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalars(xlScalars, "scalars");
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		boost::shared_ptr< Range<Scalar> > ret = XLPSession::instance().excelId(scalars);
		XL_END(ret)
	}

	XL_EXPORT OPER *xlpPythonId(OPER* xlScalars, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalars(xlScalars, "scalars");
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		boost::shared_ptr< Range<Scalar> > ret = XLPSession::instance().pythonId(scalars);
		XL_END(ret)
	}

	XL_EXPORT OPER *xlpSaveRange(OPER* xlScalars, OPER* xlId, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalars(xlScalars, "scalars");
		XlpOper xlpId(xlId, "id");
		std::string id = xlpId;
		boost::shared_ptr< Range<Scalar> > scalars = xlpScalars;
		XLSession::instance().saveRange(scalars, id);
		XL_END("success")
	}

	XL_EXPORT OPER *xlpLoadRange(OPER* xlId, OPER* xlOutput, OPER *xlTranspose, OPER* xlTrigger) {
		XL_BEGIN_OUTPUT(xlTrigger, xlTranspose, xlOutput)
		XlpOper xlpId(xlId, "name");
		std::string id = xlpId;
		boost::shared_ptr< Range<Scalar> > ret = XLSession::instance().loadRange(id);
		XL_END_OUTPUT(ret)
	}

	XL_EXPORT OPER *xlpTmpDir(OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XL_END(XLSession::instance().tmpDir())
	}

	XL_EXPORT OPER *xlpFalseTrigger(OPER* xlScalar, OPER* xlFalse, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		XlpOper xlpScalar(xlScalar, "scalar");
		XlpOper xlpFalse(xlFalse, "false");
		Scalar falsevalue;
		if (xlpFalse.missing()) {
			falsevalue = false;
		} else {
			falsevalue = Scalar(xlpFalse);
		}
		Scalar scalar = xlpScalar;
		XL_END(!(scalar.scalar() == falsevalue.scalar()))
	}

	XL_EXPORT OPER *xlpSerie(OPER* xlRef, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		static XLOPER xlBool;
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = false;
		Excel4(xlfVolatile, 0, 1, &xlBool);
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		boost::shared_ptr<Xloper> xllRef(new Xloper);
		int result = Excel4(xlCoerce, &(*xllRef), 2, xlRef, &xlInt);
		boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(0, 0));
		if (result != xlretSuccess) {
			std::cerr << "filling warning message error" << std::endl;
		} else {
			boost::shared_ptr<XLRef> ref(new XLRef(xllRef));
			ret = XLSession::instance().serie(ref);
		}
		XL_END(ret)
	}

	XL_EXPORT OPER *xlpSerieValue(OPER* xlRef, OPER* xlIndex, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		static XLOPER xlBool;
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = false;
		Excel4(xlfVolatile, 0, 1, &xlBool);
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		boost::shared_ptr<Xloper> xllRef(new Xloper);
		int result = Excel4(xlCoerce, &(*xllRef), 2, xlRef, &xlInt);
		Scalar value;
		if (result != xlretSuccess) {
			value = NullType();
			std::cerr << "filling warning message error" << std::endl;
		} else {
			boost::shared_ptr<XLRef> ref(new XLRef(xllRef));
			XlpOper xlpIndex(xlIndex, "index");
			Integer index = xlpIndex;
			value = XLSession::instance().serieValue(ref, index);
		}
		XL_END(value)
	}

	XL_EXPORT OPER *xlpVisible(OPER* xlSheet, OPER* xlSheets, OPER* xlFlag, OPER* xlTrigger) {
		XL_BEGIN(xlTrigger)
		static XLOPER xlBool;
		XlpOper xlpSheets(xlSheets, "sheets");
		XlpOper xlpSheet(xlSheet, "sheet");
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = false;
		Excel4(xlfVolatile, 0, 1, &xlBool);
		bool flag = XLSession::instance().convertBool(xlFlag, false);
		boost::shared_ptr< Range<Scalar> > ssheets = xlpSheets;
		std::vector<std::string> sheets(ssheets->size());
		for (Size i=0;i<sheets.size();++i) {
			sheets[i] = ssheets->data[i];
		}
		std::string sheet;
		try {
			sheet = xlpSheet;
		} catch(...) {
			Integer index = xlpSheet;
			REQUIRE (index<Integer(sheets.size()), "index too big: " << index);
			REQUIRE (index>=0, "index cannot be negative");
			sheet = sheets[index];
		}
		XLSession::instance().visible(sheets, sheet, flag);
		XL_END("success")
	}



}



