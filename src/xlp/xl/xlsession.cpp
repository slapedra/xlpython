/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlcellevents/all.hpp>
#include <xlp/xl/xlcommandevents/all.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlpoper.hpp>
#include <sstream>
#include <boost/tokenizer.hpp>
#include <boost/xpressive/xpressive_static.hpp>
#include <xlp/version.hpp>
#include <boost/algorithm/string.hpp> 
#include <boost/algorithm/string/trim.hpp> 

#include <iomanip>
#include <iostream>
#include <fstream>
#include <boost/archive/tmpdir.hpp>
#include <boost/archive/text_iarchive.hpp>
#include <boost/archive/text_oarchive.hpp>
#include <boost/serialization/base_object.hpp>
#include <boost/serialization/utility.hpp>
#include <boost/serialization/list.hpp>
#include <boost/serialization/assume_abstract.hpp>
#include <boost/serialization/vector.hpp>
#include <boost/serialization/variant.hpp>
#include <boost/timer.hpp>
#include <boost/filesystem.hpp>



using namespace boost::xpressive;
using namespace boost::serialization;

namespace xlp {

	namespace details {

		#define SECS_PER_DAY (60.0 * 60.0 * 24.0)

		static LPSTR menu[][2] = {
			{" &XLPython",				" "},
			{" &realtime",					" xlMenuRealtime"},
			{" &onopen",					" xlOnOpen"},
			{" -",								" "},
			{" &hide/protect",					" xlOnHideProtect"},
			{" &unhide/unprotect",					" xlOnUnhideUnprotect"},
			{" &activate/desactivate",					" xlOnActivate"},
			{" &edit mode",					" xlOnChangeMode"},
			{" -",								" "},
			{" version: "XLP_STR_PRODUCTVERSION,		" "},
			{" author: Sebastien Lapedra",		" "}
		};

		static XLOPER xMenu;
		static XLOPER xMenuList[11*2];

		XL_EXPORT bool xlOnTime(void) {
			try {
				if (XLSession::instance().activated()) {
					XLSession::instance().freeze(true, false);
					XLSession::instance().onTime();
					XLSession::instance().freeze(false);
				}
				return true;
			} catch(Error& e) {
				XLSession::instance().freeze(false);
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

		XL_EXPORT bool xlOnRecalc(void) {
			try {
				if (XLSession::instance().activated()) {
					XLSession::instance().onRecalc();
				}
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnSheet(void) {
			try {
				if (XLSession::instance().activated()) {
					XLSession::instance().onSheet();
				}
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnOpen(void) {
			try {
				if (XLSession::instance().activated()) {
					XLSession::instance().onOpen();
				}
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnWindow(void) {
			try {
				if (XLSession::instance().activated()) {
					XLSession::instance().onWindow();
				}
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnEntry(void) {
			try { 
				if (XLSession::instance().activated()) {
					XLSession::instance().onEntry();
				}
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnDel(void) {
			try {
				XLSession::instance().onDel();
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnSKey(void) {
			try {
				XLSession::instance().onSKey();
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlOnDC(void) {
			try {
				XLSession::instance().onDC();
				return true;
			} catch (Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}

		XL_EXPORT bool xlMenuRealtime(void) {
			try {
				bool flag = XLSession::instance().runningRealtime();
				XLSession::instance().realtime(!flag, XL_TIMESTEP);
				Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], 
					&details::xMenuList[2], TempBool(!flag));
				details::xlOnTime();
				return true;
			} catch(Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

		XL_EXPORT bool xlOnActivate(void) {
			try {
				bool flag = XLSession::instance().activated();
				XLSession::instance().activate(!flag);
				Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], 
					&details::xMenuList[12], TempBool(!flag));
				return true;
			} catch(Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

		XL_EXPORT bool xlOnChangeMode(void) {
			try {
				bool flag = XLSession::instance().developperMode();
				XLSession::instance().onDevelopement(!flag);
				Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], 
					&details::xMenuList[14], TempBool(!flag));
				XLOPER xlList;
				Excel4(xlfDocuments, &xlList, 0);
				XlpOper xlpList(&xlList, "list");
				boost::shared_ptr< Range<Scalar> > data = xlpList;
				for (Size i=0;i<data->size();++i) {
					std::string name = data->data[i];
				}
				return true;
			} catch(Error& e) {
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

		XL_EXPORT bool xlOnHideProtect(void) {
			try {
				XLSession::instance().freeze(true);
				XLSession::instance().onHideProtect(true);
				XLSession::instance().freeze(false);
				return true;
			} catch(Error& e) {
				XLSession::instance().freeze(false);
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

		XL_EXPORT bool xlOnUnhideUnprotect(void) {
			try {
				XLSession::instance().freeze(true);
				XLSession::instance().onHideProtect(false);
				XLSession::instance().freeze(false);
				return true;
			} catch(Error& e) {
				XLSession::instance().freeze(false);
				std::cerr << e.message() << std::endl;
				return false;
			}
		}	

	}
	
/*****************************************************************************************************************
	XLCurrentPosition
*****************************************************************************************************************/
	XLCurrentPosition::XLCurrentPosition(long id, Size row, Size col, const std::string& addr, 
		const std::string& workbook, bool transpose) 
	: transpose_(transpose), curIdSheet_(id), curRow_(row), curCol_(col), addr_(addr), workbook_(workbook) {
	}

	std::string XLCurrentPosition::idSheet() const {
		long spec  = (curIdSheet_ & 0xffff) + (curIdSheet_ & 0xff0000) + (curIdSheet_ & 0xff000000) / 0x100;
		return computeNumberRepresentation(spec);
	}

	std::string XLCurrentPosition::idPosition() const {
		long addr  = (curRow_ & 0x00ffff) * 0x000100 + (curCol_ & 0x0000ff);
		return computeNumberRepresentation(addr);
	}

	void XLCurrentPosition::addRow() const {
		if (transpose_) {
			curCol_++;
		} else {
			curRow_++;
		}
	}

	void XLCurrentPosition::addCol() const{
		if (!transpose_) {
			curCol_++;
		} else {
			curRow_++;
		}
	}

	std::string XLCurrentPosition::computeNumberRepresentation(long id) const {
		static char cTable[] = {'0','1','2','3','4','5','6','7','8','9',
			'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'};
		static Size size = 36;
		long stay = id & 0xfffff;
		std::string ret = "00000";
		Size index = 0;
		long tmp;
		long tmp2;
		while (stay) {
			tmp = stay / size;
			tmp2 = tmp*size;
			ret[index] = cTable[stay - tmp2]; 
			++index;
			stay = tmp;
		}
		return ret;
	}

	/*! XLCaller
	 */
	XLCaller::XLCaller() {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		Xloper xlCaller;
		int result = Excel4(xlfCaller, &xlCaller, 0);
		if (result != xlretSuccess) {
			FAIL("xlfCaller error");
		} 
		Xloper xlRef;
		result = Excel4(xlCoerce, &xlRef, 2, &xlCaller, &xlInt);
		if (result != xlretSuccess) {
			FAIL("ref conversion failed");
		} 
		REQUIRE (xlRef->val.mref.lpmref->count == 1, "multiple reference not supported");
		idSheet_ = xlRef->val.mref.idSheet;
		row_ = xlRef->val.mref.lpmref->reftbl[0].rwFirst;
		rows_ = xlRef->val.mref.lpmref->reftbl[0].rwLast - row_ + 1;
		col_ = xlRef->val.mref.lpmref->reftbl[0].colFirst;
		cols_ = xlRef->val.mref.lpmref->reftbl[0].colLast - col_ + 1;
		Xloper xlAddr;
		xlInt.val.w = 32;
		result = Excel4(xlfGetCell, &xlAddr, 2, &xlInt, &xlRef);
		if ( (result != xlretSuccess) || (xlAddr->xltype == xltypeErr) ) {
			FAIL("error getting address of caller");
		} 
		XlpOper xlpName(&xlAddr, "addr");
		addr_ = xlpName;
		std::vector<std::string> ret;
		boost::split(ret, addr_, boost::is_any_of("]"));
		if (ret.size()) {
			workbook_ = ret[0];
			trim_left_if(workbook_, boost::is_any_of("'"));
			trim_left_if(workbook_, boost::is_any_of("["));
		} else {
			workbook_ = addr_;
		}
	}

	/***********************************************
		XLFunction
	***********************************************/
	XLFunction::XLFunction(
		const std::string& moduleName,
		const std::string& cName, const std::string& cArgs,
		const std::string& xlName, const std::string& xlArgs,
		const std::string& desc,
		const std::string& descArgs)
	: cName_(cName), cArgs_(cArgs), xlName_(xlName), 
    xlArgs_(xlArgs), desc_(desc), 
	  moduleName_(moduleName) {	
		typedef boost::tokenizer<boost::char_separator<char> > tokenizer;
		boost::char_separator<char> sep("\n");
		std::string tmp = descArgs; 
		tmp += " ";
		tokenizer tokens(tmp, sep);
		for (tokenizer::iterator tok_iter = tokens.begin();
			tok_iter != tokens.end(); ++tok_iter) {
			descArgs_.push_back(*tok_iter);
		}
	}
		
	void XLFunction::registerFunction() {
		XLOPER xResult;
		XLOPER xDll;
		Excel(xlGetName, &xDll, 0);
		std::vector<LPXLOPER> args;
		LPXLOPER *pArgs;
		args.push_back((LPXLOPER)&xDll);
		// function code name
		args.push_back(TempStrStl(cName_.c_str()));
		// parameter codes
		args.push_back(TempStrStl(cArgs_.c_str()));
		// display name
		args.push_back(TempStrStl(xlName_.c_str()));
		// comma-delimited list of parameter names
		args.push_back(TempStrStl(xlArgs_.c_str()));
		// function type (0 = hidden, 1 = worksheet, 2 = type3)
		args.push_back(TempNum(1));
		// module name
		args.push_back(TempStrStl(moduleName_));
		// shortcut text for macro function
		args.push_back(TempStrStl(""));
		// path to help file
		args.push_back(TempStrStl(""));
		// function description
		args.push_back(TempStrStl(desc_));
		// arguments description*/
		for (Size i=0;i<descArgs_.size();i++) {
			args.push_back(TempStrStl(descArgs_[i]));
		}
		pArgs = new LPXLOPER [args.size()];
		for (Size i=0;i<args.size();i++) {
			pArgs[i] = args[i];
		}
		Integer res = Excel4v(xlfRegister, &xResult, 
			static_cast<int>(args.size()), pArgs);
		if ( (res != xlretSuccess) || (xResult.xltype == xltypeErr) ) {
			Excel4(xlFree, 0, 1, &xResult);
			Excel4(xlFree, 0, 1, &xDll);
			delete pArgs;
			FAIL ("unable to register function : " << xlName_);
		}
		delete pArgs;
		regId_ = xResult.val.num;
		Excel4(xlFree, 0, 1, &xResult);
		Excel4(xlFree, 0, 1, &xDll);
	}

	void XLFunction::unregisterFunction() {
		int result;
		if (regId_) {
			result = Excel4(xlfUnregister, 0, 1, TempNum(regId_));
			REQUIRE (result == xlretSuccess, "unable to unregister function : " << xlName_);
		}
	}

	/*******************************************************
		XLCommand
	*******************************************************/
	XLCommand::XLCommand(
		const std::string& moduleName,
		const std::string& cName, const std::string& cArgs,
		const std::string& xlName, const std::string& xlArgs,
		const std::string& desc,
		const std::string& descArgs)
	: cName_(cName), cArgs_(cArgs), xlName_(xlName), 
    xlArgs_(xlArgs), desc_(desc), 
	  moduleName_(moduleName) {	
		typedef boost::tokenizer<boost::char_separator<char> > tokenizer;
		boost::char_separator<char> sep("\n");
		std::string tmp = descArgs; 
		tmp += " ";
		tokenizer tokens(tmp, sep);
		for (tokenizer::iterator tok_iter = tokens.begin();
			tok_iter != tokens.end(); ++tok_iter) {
			descArgs_.push_back(*tok_iter);
		}
	}
		

	void XLCommand::registerCommand() {
		XLOPER xResult;
		XLOPER xDll;
		Excel(xlGetName, &xDll, 0);
		std::vector<LPXLOPER> args;
		LPXLOPER *pArgs;
		args.push_back((LPXLOPER)&xDll);
		// function code name
		args.push_back(TempStrStl(cName_.c_str()));
		// parameter codes
		args.push_back(TempStrStl(cArgs_.c_str()));
		// display name
		args.push_back(TempStrStl(xlName_.c_str()));
		// comma-delimited list of parameter names
		args.push_back(TempStrStl(xlArgs_.c_str()));
		// function type (0 = hidden, 1 = worksheet, 2 = type3)
		args.push_back(TempNum(2));
		// module name
		args.push_back(TempStrStl(moduleName_));
		// shortcut text for macro function
		args.push_back(TempStrStl(""));
		// path to help file
		args.push_back(TempStrStl(""));
		// function description
		args.push_back(TempStrStl(desc_));
		// arguments description*/
		for (Size i=0;i<descArgs_.size();i++) {
			args.push_back(TempStrStl(descArgs_[i]));
		}
		pArgs = new LPXLOPER [args.size()];
		for (Size i=0;i<args.size();i++) {
			pArgs[i] = args[i];
		}
		Integer res = Excel4v(xlfRegister, &xResult, 
			static_cast<int>(args.size()), pArgs);
		if ( (res != xlretSuccess) || (xResult.xltype == xltypeErr) ) {
			Excel4(xlFree, 0, 1, &xResult);
			Excel4(xlFree, 0, 1, &xDll);
			delete pArgs;
			FAIL ("unable to register command : " << xlName_);
		}
		delete pArgs;
		regId_ = xResult.val.num;
		Excel4(xlFree, 0, 1, &xResult);
		Excel4(xlFree, 0, 1, &xDll);
	}

	void XLCommand::unregisterCommand() {
		int result;
		if (regId_) {
			result = Excel4(xlfUnregister, 0, 1, TempNum(regId_));
			REQUIRE (result == xlretSuccess, "unable to unregister command: " << xlName_);
		}
	}

	/*******************************************************
		XLP_Session
	*******************************************************/
	XLSession::~XLSession() {
		runningRealtime_ = false;
	}

	void XLSession::registerSession() {
		registerXLCommand(
			"XLL", 
			"xlOnRecalc",
			XL_BOOL+XL_UNCALCED,
			"xlOnRecalc",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnTime",
			XL_BOOL+XL_UNCALCED,
			"xlOnTime",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnSheet",
			XL_BOOL+XL_UNCALCED,
			"xlOnSheet",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnOpen",
			XL_BOOL+XL_UNCALCED,
			"xlOnOpen",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnWindow",
			XL_BOOL+XL_UNCALCED,
			"xlOnWindow",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnEntry",
			XL_BOOL+XL_UNCALCED,
			"xlOnEntry",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnDel",
			XL_BOOL+XL_UNCALCED,
			"xlOnDel",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnSKey",
			XL_BOOL+XL_UNCALCED,
			"xlOnSKey",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnDC",
			XL_BOOL+XL_UNCALCED,
			"xlOnDC",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlMenuRealtime",
			XL_BOOL+XL_UNCALCED,
			"xlMenuRealtime",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnActivate",
			XL_BOOL+XL_UNCALCED,
			"xlOnActivate",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnChangeMode",
			XL_BOOL+XL_UNCALCED,
			"xlOnChangeMode",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnHideProtect",
			XL_BOOL+XL_UNCALCED,
			"xlOnHideProtect",
			"", 
			"", 
			"");
		registerXLCommand(
			"XLL", 
			"xlOnUnhideUnprotect",
			XL_BOOL+XL_UNCALCED,
			"xlOnUnhideUnprotect",
			"", 
			"", 
			"");
		XLOPER xDll;
		int result = xlretSuccess;
		result  = Excel4(xlGetName, &xDll, 0);
		if (result != xlretSuccess) {
			Excel4(xlFree, 0, 1, &xDll);
			FAIL ("xlGetName error");
		} else {
			XlpOper xlaId(&xDll, "id");
			id_ = xlaId;
			for (Size i=0;i<functions_.size();i++) {
				functions_[i]->registerFunction();
			}
			for (Size i=0;i<commands_.size();i++) {
				commands_[i]->registerCommand();
			}
			Excel(xlFree, 0, 1, &xDll);
			int res = Excel4(xlcOnEntry, 0, 2, 0, TempStrStl("xlOnEntry"));
			if (res != xlretSuccess) {
				FAIL("xlcOnEntry");
			}
			res = Excel4(xlcOnSheet, 0, 2, 0, TempStrStl("xlOnSheet"));
			if (res != xlretSuccess) {
				FAIL("xlcOnSheet");
			}
			res = Excel4(xlcOnWindow, 0, 2, 0, TempStrStl("xlOnWindow"));
			if (res != xlretSuccess) {
				FAIL("xlcOnWindow");
			}
			res = Excel4(xlcOnDoubleclick, 0, 2, 0, TempStrStl("xlOnDC"));
			if (res != xlretSuccess) {
				FAIL("xlcOnDC");
			}
			res = Excel4(xlcOnRecalc, 0, 2, 0, TempStrStl("xlOnRecalc"));
			if (res != xlretSuccess) {
				FAIL("xlcOnRecalc");
			}
			res = Excel4(xlcOnKey, 0, 2, TempStrStl("{DEL}"), TempStrStl("xlOnDel"));
			if (res != xlretSuccess) {
				FAIL("xlcOnKey Del");
			} 
			res = Excel4(xlcOnKey, 0, 2, TempStrStl("^~"), TempStrStl("xlOnSKey"));
			if (res != xlretSuccess) {
				FAIL("xlcOnKey Ctrl Enter");
			} 
			static XLOPER xlNow;
			int result = Excel4(xlfNow, &xlNow, 0);
			if (result != xlretSuccess) {
				FAIL("xlfNow error");
			}
			firstTime_ = xlNow.val.num;
			// menu
			details::xMenu.xltype            = xltypeMulti;
			details::xMenu.val.array.lparray = &details::xMenuList[0];
			details::xMenu.val.array.rows    = 11;
			details::xMenu.val.array.columns = 2;
			for(int i=0;i<details::xMenu.val.array.rows;++i) {
				for(int j=0;j<details::xMenu.val.array.columns;++j) {
					details::xMenuList[j+i*2].xltype  = xltypeStr;
					details::xMenuList[j+i*2].val.str = details::menu[i][j];
				}
			}
			int pos = Excel4(xlfAddMenu,0, 2, TempNum(1), (LPXLOPER)&details::xMenu);
			Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], &details::xMenuList[2], 
				TempBool(runningRealtime()));

			Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], &details::xMenuList[12], 
				TempBool(activated()));

			Excel(xlfCheckCommand,0, 4, TempNum(1), &details::xMenuList[0], &details::xMenuList[14], 
				TempBool(developperMode()));

			Excel(xlfEnableCommand,0, 4, TempNum(1), &details::xMenuList[0], &details::xMenuList[18], TempBool(false));
			Excel(xlfEnableCommand,0, 4, TempNum(1), &details::xMenuList[0], &details::xMenuList[20], TempBool(false));
		}
	}
		
	void XLSession::unregisterSession() {
		for (Size i=0;i<functions_.size();i++) {
			functions_[i]->unregisterFunction();
		}
		for (Size i=0;i<commands_.size();i++) {
			commands_[i]->unregisterCommand();
		}
		Excel(xlfDeleteMenu, 0, 2, TempNum(1), TempStrStl(" &XLPython"));
	}

	void XLSession::push_back(const boost::shared_ptr<XLFunction>& func) {
		functions_.push_back(func);
	}

	void XLSession::push_back(const boost::shared_ptr<XLCommand>& command) {
		commands_.push_back(command);
	}

	boost::shared_ptr<XLFunction> function(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs) {
		boost::shared_ptr<XLFunction> ret(new XLFunction(moduleName, cName, cArgs, 
			xlName, xlArgs, desc, descArgs));
		return ret;
	}

	boost::shared_ptr<XLCommand> command(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs) {
		boost::shared_ptr<XLCommand> ret(new XLCommand(moduleName, cName, cArgs, 
			xlName, xlArgs, desc, descArgs));
		return ret;
	}

	void registerXLFunction(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs) {
				xlp::XLSession::instance().push_back( 
					function(moduleName, cName, cArgs, 
						xlName, xlArgs, desc, descArgs));
	}

	void registerXLCommand(const std::string& moduleName,
			const std::string& cName, const std::string& cArgs,
			const std::string& xlName, const std::string& xlArgs,
			const std::string& desc,
			const std::string& descArgs) {
				xlp::XLSession::instance().push_back( 
					command(moduleName, cName, cArgs, 
						xlName, xlArgs, desc, descArgs));
	}

	void XLSession::validateTrigger(OPER* xlTrigger) {
		if (activated()) {
			XlpOper xlpTrigger(xlTrigger, "trigger");
			bool flag = true;
			try {
				flag = xlpTrigger; 
			} catch (Error&) {
			}
			if (flag == false) {
				FAIL ("invalidated function");
			}
		} else {
			FAIL ("invalidated addin");
		}
	}

	bool XLSession::convertBool(OPER* xlValue, bool defaultValue) {
		XlpOper xlpValue(xlValue, "value", true, false);
		bool value = defaultValue;
		if (!xlpValue.missing()) {
			value = xlpValue;
		}
		return value;
	}

	double XLSession::time() const {
		return  + firstTime_ + timer_.elapsed() / SECS_PER_DAY;
	}

	void XLSession::defineName(const std::string& scalar, const std::string& name) const {
		currentPosition();
		std::string id = "'"+currentPosition_->addr()+"'!"+scalar;
		static XLOPER xlBool;
		static XLOPER xlInt;
		XlpOper xlpId(id, "id");
		xlpId.setToFree(true);
		Xloper xlTmpRef;
		boost::shared_ptr<Xloper> xlRef(new Xloper());		
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = true;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		Excel4(xlfTextref, &xlTmpRef, 2, xlpId.get(), &xlBool);
		int result = Excel4(xlCoerce, &(*xlRef), 2, &xlTmpRef, &xlInt);
		if ( (result != xlretSuccess) || ( (&(*xlRef))->xltype == xltypeErr) ) {
			FAIL("error getting address of name matrix cell: " << scalar);
		} 
		XlpOper xlpMulti(&(*xlRef), "multi", false, false, true);
		boost::shared_ptr< Range<Scalar> > data = xlpMulti;
		(&(*xlRef))->val.mref.lpmref->reftbl[0].rwLast = 
			(&(*xlRef))->val.mref.lpmref->reftbl[0].rwFirst + data->rows-1;
		(&(*xlRef))->val.mref.lpmref->reftbl[0].colLast = 
			(&(*xlRef))->val.mref.lpmref->reftbl[0].colFirst + data->columns-1;
		boost::shared_ptr<XLRef> ref(new XLRef(xlRef));
		boost::shared_ptr<XLCommandEvent> com(new XLNameEvent(name, ref, this));
		recalcEvents_.push_back(com);
	}

	void XLSession::defineName(const boost::shared_ptr<XLRef>& xlRef, const std::string& name) const {
		XlpOper xlpMulti(&(*(xlRef->oper())), "multi", false, false, true);
		boost::shared_ptr< Range<Scalar> > data = xlpMulti;
		boost::shared_ptr<XLRef> ref(new XLRef(
			xlRef->row(), xlRef->col(), 
			data->rows, data->columns, 
			xlRef->idSheet()));
		boost::shared_ptr<XLCommandEvent> com(new XLNameEvent(name, ref, this));
		recalcEvents_.push_back(com);
	}

	boost::shared_ptr<XLRef> XLSession::checkOutput(const std::string& output) const {
		static XLOPER xlBool;
		static XLOPER xlInt;
		XlpOper xlpId(output, "id");
		xlpId.setToFree(true);
		boost::shared_ptr<Xloper> xlOutput(new Xloper());		
		xlBool.xltype = xltypeBool;
		xlBool.val.boolean = true;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		boost::shared_ptr<Xloper> xlRef(new Xloper());
		Excel4(xlfTextref, &(*(xlRef.get())), 2, xlpId.get(), &xlBool);
		if ((*(xlRef.get()))->xltype == xltypeRef) {
			boost::shared_ptr<XLRef> ref(new XLRef(xlRef));
			return ref;
		} else {
			int result = Excel4(xlCoerce, &(*xlOutput), 2, &(*(xlRef.get())), &xlInt);
			if ( (result != xlretSuccess) || ( (&(*xlOutput))->xltype == xltypeErr) ) {
				FAIL("error getting address of name matrix cell: " << output);
			} 
			boost::shared_ptr<XLRef> ref(new XLRef(xlOutput));
			return ref;
		}
	}

	std::string XLSession::fillOutput(const boost::shared_ptr< Range<Scalar> >& data) const {
		if ( (!output_.empty()) && (xlref_.get() != NULL)) {
			boost::shared_ptr<XLCommandEvent> com(new XLOutputEvent(data, transpose_, xlref_, this));
			recalcEvents_.push_back(com);
			if (!transpose_) { 
				if ( (xlref_->cols() < data->columns) || (xlref_->rows() < data->rows)) {
					return "more";
				}
			} else {
				if ( (xlref_->cols() < data->rows) || (xlref_->rows() < data->columns)) {
					return "more";
				}
			}
			return "success";
		} else {
			return "";
		}
	}

	std::string XLSession::fillOutput(const boost::shared_ptr<XLRef>& xlref, 
		const boost::shared_ptr< Range<Scalar> >& data, bool transpose, Real time) const {
		if ((xlref.get() != NULL)) {
			boost::shared_ptr<XLCommandEvent> com(new XLOutputEvent(data, transpose, xlref, this, time));
			recalcEvents_.push_back(com);
			if (!transpose) { 
				if ( (xlref->cols() < data->columns) || (xlref->rows() < data->rows)) {
					return "more";
				}
			} else {
				if ( (xlref->cols() < data->rows) || (xlref->rows() < data->columns)) {
					return "more";
				}
			}
			return "success";
		} else {
			return "";
		}
	}

	void XLSession::onTime() const {
		if (runningRealtime_) {
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			xlInt.val.w = 76;
			Xloper xlPath;
			Excel4(xlfGetDocument, &xlPath, 1, &xlInt);
			XlpOper xlAddr(&xlPath, "addr");
			std::string addr = xlAddr;
			std::vector<std::string> ret;
			boost::split(ret, addr, boost::is_any_of("]"));
			trim_if(ret[0], boost::is_any_of("'"));
			trim_if(ret[0], boost::is_any_of("["));
			addr = "'"+ret[0]+"'";
			boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, std::list< boost::shared_ptr<XLRef> >());
			if (cellmap.get() != NULL) {
				freeze(true);
				XLCellEventMap::const_iterator iter = cellmap->begin();
				while (iter != cellmap->end()) {
					if(!iter->handler->onTime()) {
					}
					++iter;
				}
				freeze(false);
			}
			static XlpOper cmd("xlOnTime");
			double nextTime = time();
			static XLOPER xlNow;
			xlNow.xltype = xltypeNum;
			xlNow.val.num = nextTime + XLSession::instance().timestep() / SECS_PER_DAY;
			int result = Excel4(xlcOnTime, 0, 2, &xlNow, cmd.get());
			if (result != xlretSuccess) {
				std::cerr << "xlcOnTime failed" << std::endl;
			}
		}
		FreeAllTempMemory();
	}

	void XLSession::onOpen() const {
		workbooks_.erase(workbooks_.begin(), workbooks_.end());
		onWindow();
	}

	void XLSession::onWindow() const {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 38;
		Xloper xlSheet;
		Excel4(xlfGetWorkbook, &xlSheet, 1, &xlInt);

		XLOPER xlList;
		int result = Excel4(xlfDocuments, &xlList, 0);
		if (result != xlretSuccess) {
			std::cerr << "error in getting workbook list" << std::endl;
		}
		XlpOper xlpList(&xlList, "list");
		boost::shared_ptr< Range<Scalar> > data = xlpList;
		std::set<std::string> workbooks;
		for (Size i=0;i<data->size();++i) {
			std::string name = data->data[i];
			workbooks.insert(name);
		}
		std::set<std::string>::iterator iter;
		std::set<std::string>::iterator iter2;
		iter = workbooks.begin();
		while (iter != workbooks.end()) {
			iter2 = workbooks_.find(*iter);
			if (iter2 == workbooks_.end()) {
				std::string addr = "'"+std::string(*iter)+"'";
				worksheet_ = addr;
				workbooks_.insert(*iter);
				boost::shared_ptr<XLCellEventMap> cellmap = 
					cellEvents(addr, std::list< boost::shared_ptr<XLRef> >(), true);
				try {
					if (cellmap.get() != NULL) {
						freeze(true);
						XLCellEventMap::const_iterator iter = cellmap->begin();
						while (iter != cellmap->end()) {
							if (iter->handler->onOpen()) {
							}
							++iter;
						}
						freeze(false);
					}
				} catch (...) {
				}
				
			}
			++iter;
		}
		iter = workbooks_.begin();
		while (iter != workbooks_.end()) {
			iter2 = workbooks.find(*iter);
			if (iter2 == workbooks.end()) {
				if (onWorkbookCloseCallback_) {
					onWorkbookCloseCallback_(*iter);
				}
				SerieValue::erase(*iter);
				std::map<std::string, boost::shared_ptr<XLCellEventMap> >::iterator iter3 = cellmaps_.find(*iter);
				if (iter3 != cellmaps_.end()) {
					cellmaps_.erase(iter3);
				}
				iter = workbooks_.erase(iter);
			} else {
				++iter;
			}
		}
		FreeAllTempMemory();
	}

	void XLSession::onSheet() const {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 76;
		Xloper xlPath;
		Xloper xlId;
		Excel4(xlfGetDocument, &xlPath, 1, &xlInt);
		Excel4(xlSheetId, &xlId, 1, &xlPath);
		XlpOper xlAddr(&xlPath, "addr");
		std::string addr = xlAddr;
		std::vector<std::string> ret;
		boost::split(ret, addr, boost::is_any_of("]"));
		trim_if(ret[0], boost::is_any_of("'"));
		trim_if(ret[0], boost::is_any_of("["));
		addr = "'"+ret[0]+"'";
		XLEvent::idSheet(xlId->val.mref.idSheet);
		boost::shared_ptr<XLCellEventMap> cellmap = 
			cellEvents(addr, std::list< boost::shared_ptr<XLRef> >());
		if (cellmap.get() != NULL) {
			XLCellEventMap::const_iterator iter = cellmap->begin();
			while (iter != cellmap->end()) {
				if (iter->handler->onSheet()) {
				}
				++iter;
			}
		}
		FreeAllTempMemory();
	}

	boost::shared_ptr<XLCellEventMap> XLSession::cellEvents(const std::string& src,
		const std::list< boost::shared_ptr<XLRef> >& xlRefs, bool flag) const {
		XLCellEvent::time(true);
		XLEvent::time(true);
		if (!flag && !developpermode_) {
			std::map<std::string, boost::shared_ptr<XLCellEventMap> >::iterator iter = cellmaps_.find(src);
			if (iter != cellmaps_.end()) {
				return iter->second;
			} else {
				return boost::shared_ptr<XLCellEventMap>(new XLCellEventMap());
			}
		} else {
			boost::shared_ptr<XLCellEventMap> cellmap(new XLCellEventMap());
			try {
				Xloper xlNames;
				int result;
				result = Excel4(xlfNames, &xlNames, 0);
				if (result != xlretSuccess) {
					std::cerr << "error in getting names list" << std::endl;;
				} else {
					XlpOper xlpNames(&xlNames, "names");
					boost::shared_ptr< Range<Scalar> > names = xlpNames;
					for (Size i=0;i<names->size();++i) {
						std::string name = names->data[i];
						try {
							boost::shared_ptr<XLRef> xlRef = checkOutput(src+"!"+name);
							bool flag = false;
							if (!xlRefs.empty()) {
								std::list< boost::shared_ptr<XLRef> >::const_iterator iter = xlRefs.begin();
								while (iter!= xlRefs.end()) {
									if ((*iter)->inter(xlRef)) {
										flag = true;
										break;
									}
									++iter;
								}
							} else {
								flag = true;
							}
							if (flag) {
								try {
									if (CustomXLCellEvent::populate(xlRef, name, nameKey_, src, cellmap, this)) {
									}
								} catch (...) {
								}
							}
						} catch (...) {
						}
					}
				}
			} catch (...) {
			}
			std::map<std::string, boost::shared_ptr<XLCellEventMap> >::iterator iter = cellmaps_.find(src);
			if (iter == cellmaps_.end()) {
				cellmaps_.insert(std::make_pair(src, cellmap));
			} else {
				cellmaps_[src] = cellmap;
			}
			return cellmap;
		}
	}

	void XLSession::onEntry() const {
		boost::shared_ptr<Xloper> xlActiveCell(new Xloper());
		Xloper xlRef;
		Xloper xlName;
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		int result = Excel4(xlfSelection, &(*(xlActiveCell.get())), 0);
		if (result != xlretSuccess) {
			std::cerr << "xlfSelection error" << std::endl;
			FreeAllTempMemory();
			return;
		}
		boost::shared_ptr<XLRef> xlActiveRef(new XLRef(xlActiveCell));
		result = Excel4(xlfReftext, &xlRef, 1, &(*(xlActiveCell.get())));
		if (result != xlretSuccess) {
			std::cerr << "xlfRefText error" << std::endl;
			FreeAllTempMemory();
			return;
		}
		XlpOper xlAddr(&xlRef, "addr");
		std::string addr = xlAddr;
		std::vector<std::string> ret;
		boost::split(ret, addr, boost::is_any_of("]"));
		trim_if(ret[0], boost::is_any_of("'"));
		trim_if(ret[0], boost::is_any_of("["));
		addr = "'"+ret[0]+"'";
		boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, xlActiveRef);
		bool flag = false;
		if (cellmap.get() != NULL) {
			Size row = xlActiveRef->row();
			Size col = xlActiveRef->col();
			for (Size i=0;i<xlActiveRef->rows();++i) {
				for (Size j=0;j<xlActiveRef->cols();++j) {
					boost::shared_ptr<XLCellEvent> handler = cellmap->handler(row+i, col+j, xlActiveRef->idSheet());
					if (!(handler.get() == NULL)) {
						if (!flag) {
							freeze(true, false);
							flag = true;
						}
						if (!handler->onEntry()) {
						}
					}
				}
			}
			if (flag) {
				freeze(false);
			}
		}
		FreeAllTempMemory();
	}

	void XLSession::onDel() const {
		boost::shared_ptr<Xloper> xlActiveCell(new Xloper());
		Xloper xlRef;
		Xloper xlName;
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		int result = Excel4(xlfSelection, &(*(xlActiveCell.get())), 0);
		if (result != xlretSuccess) {
			std::cerr << "xlfSelection error" << std::endl;
			FreeAllTempMemory();
			return;
		}
		if ((*(xlActiveCell.get()))->xltype != xltypeStr) {
			freeze(true, false);
			Excel4(xlSet, 0, 1, &(*(xlActiveCell.get())));
			if (XLSession::instance().activated()) {
				boost::shared_ptr<XLRef> xlActiveRef(new XLRef(xlActiveCell));
				result = Excel4(xlfReftext, &xlRef, 1, &(*(xlActiveCell.get())));
				if (result != xlretSuccess) {
					std::cerr << "xlfRefText error" << std::endl;
					FreeAllTempMemory();
					return;
				}
				XlpOper xlAddr(&xlRef, "addr");
				std::string addr = xlAddr;
				std::vector<std::string> ret;
				boost::split(ret, addr, boost::is_any_of("]"));
				trim_if(ret[0], boost::is_any_of("'"));
				trim_if(ret[0], boost::is_any_of("["));
				addr = "'"+ret[0]+"'";
				boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, xlActiveRef);
				if (cellmap.get() != NULL) {
					Size row = xlActiveRef->row();
					Size col = xlActiveRef->col();
					for (Size i=0;i<xlActiveRef->rows();++i) {
						for (Size j=0;j<xlActiveRef->cols();++j) {
							boost::shared_ptr<XLCellEvent> handler = cellmap->handler(row+i, col+j, xlActiveRef->idSheet());
							if (!(handler.get() == NULL)) {
								if (!handler->onEntry()) {
								}
							}
						}
					}
				}
				FreeAllTempMemory();
				freeze(false);
			}
		}
	}

	void XLSession::onSKey() const {
bool flag = false;
		if (XLSession::instance().activated()) {
			boost::shared_ptr<Xloper> xlActiveCell(new Xloper());
			Xloper xlRef;
			Xloper xlName;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			xlInt.val.w = xltypeRef;
			int result = Excel4(xlfSelection, &(*(xlActiveCell.get())), 0);
			if (result != xlretSuccess) {
				std::cerr << "xlfSelection error" << std::endl;
				FreeAllTempMemory();
				return;
			}
			if ((*(xlActiveCell.get()))->xltype != xltypeStr) {
				boost::shared_ptr<XLRef> xlActiveRef(new XLRef(xlActiveCell));
				result = Excel4(xlfReftext, &xlRef, 1, &(*(xlActiveCell.get())));
				if (result != xlretSuccess) {
					std::cerr << "xlfRefText error" << std::endl;
					XLSession::instance().freeze(false);
					return;
				}
				XlpOper xlAddr(&xlRef, "addr");
				std::string addr = xlAddr;
				std::vector<std::string> ret;
				boost::split(ret, addr, boost::is_any_of("]"));
				trim_if(ret[0], boost::is_any_of("'"));
				trim_if(ret[0], boost::is_any_of("["));
				addr = "'"+ret[0]+"'";
				boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, xlActiveRef);
				bool fflag = false;
				if (cellmap.get() != NULL) {
					Size row = xlActiveRef->row();
					Size col = xlActiveRef->col();
					for (Size i=0;i<xlActiveRef->rows();++i) {
						for (Size j=0;j<xlActiveRef->cols();++j) {
							boost::shared_ptr<XLCellEvent> handler = cellmap->handler(row+i, col+j, xlActiveRef->idSheet());
							if (handler.get() != NULL) {
								if (!fflag) {
									freeze(true, false);
									fflag = true;
								}
								if (!handler->onSKey()) {
									flag = true;
								}
							} else {
								flag = true;
							}
						}
					}
					if (flag) {
						freeze(false);
					}
				} else {
					flag = true;
				}
				freeze(false);
				FreeAllTempMemory();
			} else {
				flag = true;
			}
			if (flag) {
				Excel4(xlcSendKeys, 0, 1, TempStrStl("^{ENTER}"));
			}
		} else {
			Excel4(xlcSendKeys, 0, 1, TempStrStl("^{ENTER}"));
		}
	}

	void XLSession::onDC() const {
		bool flag = false;
		if (XLSession::instance().activated()) {
			boost::shared_ptr<Xloper> xlActiveCell(new Xloper());
			Xloper xlRef;
			Xloper xlName;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			xlInt.val.w = xltypeRef;
			int result = Excel4(xlfSelection, &(*(xlActiveCell.get())), 0);
			if (result != xlretSuccess) {
				std::cerr << "xlfSelection error" << std::endl;
				FreeAllTempMemory();
				return;
			}
			if ((*(xlActiveCell.get()))->xltype != xltypeStr) {
				XLSession::instance().freeze(true, false);
				boost::shared_ptr<XLRef> xlActiveRef(new XLRef(xlActiveCell));
				result = Excel4(xlfReftext, &xlRef, 1, &(*(xlActiveCell.get())));
				if (result != xlretSuccess) {
					std::cerr << "xlfRefText error" << std::endl;
					XLSession::instance().freeze(false);
					return;
				}
				XlpOper xlAddr(&xlRef, "addr");
				std::string addr = xlAddr;
				std::vector<std::string> ret;
				boost::split(ret, addr, boost::is_any_of("]"));
				trim_if(ret[0], boost::is_any_of("'"));
				trim_if(ret[0], boost::is_any_of("["));
				addr = "'"+ret[0]+"'";
				boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, xlActiveRef);
				bool fflag = false;
				if (cellmap.get() != NULL) {
					Size row = xlActiveRef->row();
					Size col = xlActiveRef->col();
					for (Size i=0;i<xlActiveRef->rows();++i) {
						for (Size j=0;j<xlActiveRef->cols();++j) {
							boost::shared_ptr<XLCellEvent> handler = cellmap->handler(row+i, col+j, xlActiveRef->idSheet());
							if (handler.get() != NULL) {
								if (!fflag) {
									freeze(true, false);
									fflag = true;
								}
								if (!handler->onDC()) {
									flag = true;
								}
							} else {
								flag = true;
							}
						}
					}
					if (fflag) {
						freeze(false);
					}
				} else {
					flag = true;
				}
				XLSession::instance().freeze(false);
				FreeAllTempMemory();
			} else {
				flag = true;
			}
			if (flag) {
				Excel4(xlcSendKeys, 0, 1, TempStrStl("{F2}"));
			}
		} else {
			Excel4(xlcSendKeys, 0, 1, TempStrStl("{F2}"));
		}
	}

	void XLSession::onRecalc() const {
		static XLOPER xlNow;
		while (recalcEvents_.size() != 0) {
			freeze(true);
			std::list< boost::shared_ptr<XLRef> > xlRefs;
			std::list< boost::shared_ptr<XLCommandEvent> >::iterator iter = recalcEvents_.begin();
			boost::shared_ptr<XLRef> xlRef;
			while (iter != recalcEvents_.end()) {
				if ( (*iter)->execute(xlRef) ) {
					break;
				}
				++iter; 
				if ( (xlRef.get()) != NULL ) {
					xlRefs.push_back(xlRef);
				}
			}
			recalcEvents_.erase(recalcEvents_.begin(), recalcEvents_.end());
			std::list< boost::shared_ptr<XLCellEvent> > handlers;
			static XLOPER xlInt;
			xlInt.xltype = xltypeInt;
			xlInt.val.w = 88;
			Xloper xlPath;
			Excel4(xlfGetDocument, &xlPath, 1, &xlInt);
			XlpOper xlpPath(&xlPath, "path");
			std::string addr = xlpPath;
			addr = "'"+addr+"'";
			boost::shared_ptr<XLCellEventMap> cellmap = cellEvents(addr, xlRefs);
			if (cellmap.get() != NULL) {
				std::list< boost::shared_ptr<XLRef> >::iterator iter2 = xlRefs.begin();
				while (iter2 != xlRefs.end()) {
					Size row = (*iter2)->row();
					Size col = (*iter2)->col();
					for (Size i=0;i<(*iter2)->rows();++i) {
						for (Size j=0;j<(*iter2)->cols();++j) {
							boost::shared_ptr<XLCellEvent> handler = cellmap->handler(row+i, col+j, (*iter2)->idSheet());
							if (!(handler.get() == NULL)) {
								if (!handler->onEntry()) {
								}
								handlers.push_back(handler);
							}
						}
					}
					++iter2;
				}
			};
			freeze(false);
		}
		FreeAllTempMemory();
	}

	void XLSession::freeze(bool newCalcSettingFlag, bool display) const {
		static bool oldCalcSettingFlag = false;
		static XLOPER xlOldCalcSetting;
		if (newCalcSettingFlag) {
			if (display) {
				static XLOPER xlBool;
				xlBool.xltype = xltypeBool;
				xlBool.val.boolean = false;
				int result = Excel4(xlcEcho, 0, 1, &xlBool);
				if (result != xlretSuccess) {
					FAIL("error in freezing screen update");
				}
			}
			if (oldCalcSettingFlag != newCalcSettingFlag) {
				static XLOPER xlArg;
				xlArg.xltype = xltypeInt;
				xlArg.val.w = 14;
				int result = Excel4(xlfGetDocument, &xlOldCalcSetting, 1, &xlArg);
				if (result != xlretSuccess) {
					FAIL("error in old calculation setting");
				}
				oldCalcSettingFlag = newCalcSettingFlag;
			} 
			static XLOPER xlCalcSetting;
			xlCalcSetting.xltype = xltypeNum;
			xlCalcSetting.val.num = 3;
			int result = Excel4(xlcCalculation, 0, 1, &xlCalcSetting);
			if (result != xlretSuccess) {
				FAIL("error in changing calculation setting");
			}
		} else {
			if (oldCalcSettingFlag) {
				int result = Excel4(xlcCalculation, 0, 1, &xlOldCalcSetting);
				if (result != xlretSuccess) {
					FAIL("error in setting old calculation setting");
				}
				oldCalcSettingFlag = false;
			}
		}
	}
	
	void XLSession::realtime(bool flag, Real step) const { 
		runningRealtime_ = flag;
		timestep_ = step;
		if (timestep_ < XL_TIMESTEP) {
			timestep_ = XL_TIMESTEP;
		}
	}

	void XLSession::enterFunction(OPER* xlTrigger, OPER* xlTranspose, OPER* xlOutput) const {
		flagCurrentPosition_ = false;
		XLSession::instance().validateTrigger(xlTrigger);
		if (xlTranspose != NULL) {
			transpose_ = XLSession::instance().convertBool(xlTranspose, false);
		} else {
			transpose_ = false;
		}
		if (xlOutput != NULL) {
			XlpOper xlpOutput(xlOutput, "output");
			if (!xlpOutput.missing()) {
				output_ = xlpOutput;
				XLSession::instance().currentPosition();
			} else {
				output_ = "";
			}
		} else {
			output_ = "";
		}
	}

	void XLSession::exitFunction() const {
		flagCurrentPosition_ = false;
		transpose_ = false;
		output_ = "";
	}

	boost::shared_ptr<XLCurrentPosition> XLSession::currentPosition() const {
		if (!flagCurrentPosition_) {
			XLCaller caller;
			if (output_.empty()) {
				currentPosition_.reset(
					new XLCurrentPosition(caller.idSheet(), caller.row(), caller.col(), caller.addr(), 
						caller.workbook(), transpose_));
			} else {
				std::string id = "'"+caller.addr()+"'!"+output_; 
				xlref_ = XLSession::instance().checkOutput(id);
				currentPosition_.reset(
					new XLCurrentPosition(xlref_->idSheet(), xlref_->row(), xlref_->col(), caller.addr(), 
					caller.workbook(), transpose_));
			}
			flagCurrentPosition_ = true;
		}
		return currentPosition_;
	}

	boost::shared_ptr<XLRef> XLSession::currentRef(const XLCaller& caller, const std::string& output) const {
		REQUIRE(!output.empty(), "empty ouput");
		boost::shared_ptr<XLRef> ref;
		std::string id = "'"+caller.addr()+"'!"+output; 
		return XLSession::instance().checkOutput(id);
	}

	std::string XLSession::workbookPath() const {
		XLCaller caller;
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 2;
		XlpOper xlpName(caller.addr());
		Xloper xlPath;
		int result = Excel4(xlfGetDocument, &xlPath, 2, &xlInt, xlpName.get());
		if ( (result != xlretSuccess) || (xlPath->xltype == xltypeErr) ) {
			FAIL("error getting path of the workbook");
		} 
		XlpOper xlpPath(&xlPath, "path");
		return std::string(xlpPath);
	}

	std::string XLSession::workbookName() const {
		XLCaller caller;
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 88;
		XlpOper xlpName(caller.addr());
		Xloper xlPath;
		int result = Excel4(xlfGetDocument, &xlPath, 2, &xlInt, xlpName.get());
		if ( (result != xlretSuccess) || (xlPath->xltype == xltypeErr) ) {
			FAIL("error getting name of the workbook");
		} 
		XlpOper xlpPath(&xlPath, "path");
		return std::string(xlpPath);
	}

	std::string XLSession::worksheetName(bool little) const {
		XLCaller caller;
		if (little) {
			std::vector<std::string> ret;
			boost::split(ret, caller.addr(), boost::is_any_of("]"));
			return ret[ret.size()-1];
		} else {
			return caller.addr();
		}
	}

	std::string XLSession::formula(const boost::shared_ptr< Range<Scalar> >& data, bool transpose, 
		const std::string& scalar, Real time) const {
		currentPosition();
		std::string id = "'"+currentPosition_->addr()+"'!"+scalar;
		boost::shared_ptr<XLRef> xlRef = checkOutput(id);
		return formula(xlRef, data, transpose, time);
	}

	std::string XLSession::formula(const boost::shared_ptr<XLRef>& xlref, 
		const boost::shared_ptr< Range<Scalar> >& data, bool transpose, Real time) const {
		if (xlref.get() != NULL) {
			boost::shared_ptr<XLCommandEvent> com(new XLFormulaEvent(data, transpose, xlref, this, time));
			recalcEvents_.push_back(com);
			if (!transpose) { 
				if ( (xlref->cols() < data->columns) || (xlref->rows() < data->rows)) {
					return "more";
				}
			} else {
				if ( (xlref->cols() < data->rows) || (xlref->rows() < data->columns)) {
					return "more";
				}
			}
			return "success";
		} else {
			return "";
		}
	}

	template<class Archive>
	void serialize(Archive & ar, NullType& type, const unsigned int version) {
	}

	template<class Archive>
	void serialize(Archive & ar, NAType& type, const unsigned int version) {
	}

	template<class Archive>
	void serialize(Archive & ar, Scalar& scalar, const unsigned int version) {
		ScalarDef sd = scalar.scalar();
		ar & sd;
		scalar = Scalar(sd);
	}

	template<class Archive>
	void serialize(Archive & ar, boost::shared_ptr< Range<Scalar> >& range, const unsigned int version) {
    ar & range->rows;
    ar & range->columns;
		ar & range->data;
	}

	bool XLSession::isProtected(const std::string& sheet) const {
		static XLOPER xlOldProtection;
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = 7;
		if (sheet.empty()) {
			Excel4(xlfGetDocument, &xlOldProtection, 1, &xlInt);
		} else {
			Excel4(xlfGetDocument, &xlOldProtection, 2, &xlInt, TempStrStl(sheet));
		}
		XlpOper xlpFlag(&xlOldProtection, "");
		bool flag = xlpFlag;
		return flag;
	}

	std::string XLSession::tmpDir() const {
		std::string id = std::string(boost::archive::tmpdir())+"/xlpython";
		boost::filesystem::create_directory(id);
		return id;
	}

	void XLSession::saveRange(const boost::shared_ptr< Range<Scalar> >& range, const std::string& id) const {
		std::string filename = std::string(tmpDir())+"/"+id+".xlpar";
		std::ofstream ofs(filename.c_str());
    boost::archive::text_oarchive oa(ofs);
    oa << range;
	}
	
	boost::shared_ptr< Range<Scalar> > XLSession::loadRange(const std::string& id) const {
		boost::shared_ptr< Range<Scalar> > range(new Range<Scalar>(0, 0));
		std::string filename = std::string(tmpDir())+"/"+id+".xlpar";
		std::ifstream ifs(filename.c_str());
    boost::archive::text_iarchive ia(ifs);
    ia >> range;
		return range;
	}

	void XLSession::message(const std::string& str) const {
		boost::shared_ptr<XLCommandEvent> com(new XLMessageEvent(str, this));
		recalcEvents_.push_back(com); 
	}

	void XLSession::error(const std::string& str) const { 
		boost::shared_ptr<XLCommandEvent> com(new XLAlertEvent(str, this, false));
		recalcEvents_.push_back(com); 
	}
	
	void XLSession::warning(const std::string& str) const { 
		boost::shared_ptr<XLCommandEvent> com(new XLAlertEvent(str, this, true));
		recalcEvents_.push_back(com); 
	}

	void XLSession::unprotect(bool flag) const {
		static XLOPER xlBool;
		static bool oldFlag = isProtected();
		bool newFlag = isProtected();
		xlBool.xltype = xltypeBool;
		int result;
		if (flag) {
			oldFlag = isProtected();
			xlBool.val.boolean = false;
			result = Excel4(xlcProtectDocument, 0, 1, &xlBool);
			if (result != xlretSuccess) {
				FAIL("xlcProtectDocument error");
			}
		} else {
			xlBool.val.boolean = oldFlag;
			result = Excel4(xlcProtectDocument, 0, 1, &xlBool);
			if (result != xlretSuccess) {
				FAIL("xlcProtectDocument error");
			}
		}
	}

	void XLSession::visible(const std::vector<std::string>& sheets, const std::string& sheet, bool activate) const {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;

		// active sheet
		xlInt.val.w = 38;
		Xloper xlSheet;
		Excel4(xlfGetWorkbook, &xlSheet, 1, &xlInt);
		XlpOper xlpSheet(&xlSheet, "sheets");
		std::string activesheet = xlpSheet;
		
		boost::shared_ptr<XLCommandEvent> com(
			new XLVisibleEvent(sheets, sheet, activesheet, activate));
		recalcEvents_.push_back(com);
	}

	void XLSession::onHideProtect(bool flag) const {
		static XLOPER xlInt;
		static XLOPER xlBool;
		xlInt.xltype = xltypeInt;
		xlBool.xltype = xltypeBool;
		// active sheet
		xlInt.val.w = 38;
		Xloper xlSheet;
		Excel4(xlfGetWorkbook, &xlSheet, 1, &xlInt);
		XlpOper xlpSheet(&xlSheet, "sheets");
		std::string activesheet = xlpSheet;
		// sheets
		xlInt.val.w = 1;
		Xloper xlSheets;
		Excel4(xlfGetWorkbook, &xlSheets, 1, &xlInt);
		XlpOper xlpSheets(&xlSheets, "sheets");
		boost::shared_ptr< Range<Scalar> > sheets = xlpSheets;
		// src
		xlInt.val.w = 88;
		Xloper xlPath;
		Excel4(xlfGetDocument, &xlPath, 1, &xlInt);
		XlpOper xlpPath(&xlPath, "path");
		std::string src = xlpPath;
		// unprotect workbook
		xlBool.val.boolean = false;
		Excel4(xlcWorkbookProtect, 0, 1, &xlBool);
		// unprotect and hide
		for (Size i=0;i<sheets->size();++i) {
			std::string sheet = sheets->data[i];
			std::vector<std::string> ret;
			boost::split(ret, sheet, boost::is_any_of("]"));
			sheet = ret[1];
			Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(sheet));
			// unprotect
			xlBool.val.boolean = false;
			Excel4(xlcProtectDocument, 0, 1, &xlBool);
			// unhide
			Excel4(xlcWorkbookUnhide, 0, 0);
		}
		// cleaning names;
		Xloper xlNames;
		int result;
		result = Excel4(xlfNames, &xlNames, 0);
		if (result != xlretSuccess) {
			std::cerr << "error in getting names list" << std::endl;;
		} else {
			XlpOper xlpNames(&xlNames, "names");
			boost::shared_ptr< Range<Scalar> > names = xlpNames;
			for (Size j=0;j<names->size();++j) {
				std::string name = names->data[j];
				try {
					boost::shared_ptr<XLRef> xlRef = checkOutput("'"+src+"'!"+name);
				} catch (...) {
					Excel4(xlcDeleteName, 0, 1, TempStrStl(name));
				}
			}
		}
		// hide / unhide columns rows
		for (Size i=0;i<sheets->size();++i) {
			std::string sheet = sheets->data[i];
			std::vector<std::string> ret;
			boost::split(ret, sheet, boost::is_any_of("]"));
			sheet = ret[1];
			Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(sheet));
			boost::shared_ptr<XLRef> xlRefA1 = checkOutput("'["+src+"]"+sheet+"'!A1");
			Xloper xlNames;
			int result;
			result = Excel4(xlfNames, &xlNames, 0);
			if (result != xlretSuccess) {
				std::cerr << "error in getting names list" << std::endl;;
			} else {
				XlpOper xlpNames(&xlNames, "names");
				boost::shared_ptr< Range<Scalar> > names = xlpNames;
				for (Size j=0;j<names->size();++j) {
					std::string name = names->data[j];
					static sregex regstr = as_xpr(nameKey_+"_code_") >> (s1 = +_);
					smatch what;
					if (flag) {
						xlInt.val.w = 1;
					} else {
						xlInt.val.w = 2;
					}
					if (regex_match(name, what, regstr)) {
						std::string key = what[1]; 
						boost::shared_ptr<XLRef> xlRef = checkOutput("'"+src+"'!"+name);
						if (xlRef->idSheet() == xlRefA1->idSheet()) {
							Excel4(xlcColumnWidth, 0, 4, NULL, &(*(xlRef->oper())), NULL, &xlInt);
						}
					}
				}
			}
		}
		// unprotect or protect
		for (Size i=0;i<sheets->size();++i) {
			std::string sheet = sheets->data[i];
			std::vector<std::string> ret;
			boost::split(ret, sheet, boost::is_any_of("]"));
			sheet = ret[1];
			xlBool.val.boolean = flag;
			Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(sheet));
			// protect
			Excel4(xlcProtectDocument, 0, 1, &xlBool);
			// hide
			static sregex regstr = as_xpr(nameKey_+"_") >> (s1 = +_);
			smatch what;
			if (regex_match(sheet, what, regstr)) {
				std::string key = what[1];
				if (flag) {
					Excel4(xlcWorkbookHide, 0, 0);
				} else {
					Excel4(xlcWorkbookUnhide, 0, 0);
				}
			}
		}
		xlBool.val.boolean = flag;
		Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(activesheet));
		Excel4(xlcWorkbookProtect, 0, 1, &xlBool);
		FreeAllTempMemory();
	}

	Scalar XLSession::serieValue(const boost::shared_ptr<XLRef>& xlRef, Size index) const {
		Scalar ret;
		SerieValue serie(xlRef);
		ret = serie.load(index);
		if ( ret.missing() || ret.error()) {
			UserSerieValue serie(xlRef);
			return serie.load(index);
		} 
		return ret;
	}

	boost::shared_ptr< Range<Scalar> > XLSession::serie(const boost::shared_ptr<XLRef>& xlRef) const {
		boost::shared_ptr< Range<Scalar> > ret;
		SerieValue serie(xlRef);
		ret = serie.range();
		if (ret->size() == 0) {
				UserSerieValue serie(xlRef);
				return serie.range();
		}
		return ret;
	}

	void XLSession::calculateNow() const {
		Excel4(xlcCalculateNow, 0, 0);
	}
}