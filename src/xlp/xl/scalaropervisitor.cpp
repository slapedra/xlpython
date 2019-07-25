/*
 2008 Sebastien Lapedra
*/


#include <xlp/config.hpp>
#include <xlp/xl/scalaropervisitor.hpp>
#include <boost/lexical_cast.hpp>
#include <iostream>
#include <sstream>
#include <stdlib.h>
#include <algorithm>
#include <math.h>
 
namespace xlp {

	#define MIN(X, Y)  ((X) < (Y) ? (X) : (Y))

	std::string toExcelString(const std::string& str) {
		#if PY_VERSION_HEX >= 0x03000000
		PyObject *pobj = PyUnicode_FromString(str.c_str());
		PyObject *pobj2 = PyUnicode_AsLatin1String(pobj);
		char *cstr = PyBytes_AsString(pobj2);
		std::string ret = std::string(cstr);
		delete cstr;
		return ret;
		#else
		return str;
		#endif
	}

	/************************************************************************
		ScalarToOperVisitor
	************************************************************************/
	void ScalarToOperVisitor::operator()(const Integer &i) const {
		oper_->xltype = xltypeNum;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
    oper_->val.num = i;
	}
	
	void  ScalarToOperVisitor::operator()(const Real &r) const {
		oper_->xltype = xltypeNum;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
   	oper_->val.num = r;
	}
    
	void ScalarToOperVisitor::operator()(const bool &b) const {
		oper_->xltype = xltypeBool;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
	  oper_->val.boolean = b;
	}

	void ScalarToOperVisitor::operator()(const std::string& ss) const {
		std::string s = toExcelString(ss);
		Size len = Size(MIN(XL_MAX_STR_LEN - 1, s.length()));
		oper_->val.str = new char[len + 1];
		oper_->xltype = xltypeStr;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
		oper_->val.str[0] = len;
		if (len) {
			std::copy(s.c_str(), s.c_str()+len, oper_->val.str + 1);
		}
	}
		
	void ScalarToOperVisitor::operator()(const NullType&) const {
		oper_->xltype = xltypeMissing;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
	}

	void ScalarToOperVisitor::operator()(const NAType&) const {
		oper_->xltype = xltypeErr;
		oper_->val.err = xlerrNA;
		if (dllfree_)
			oper_->xltype |= xlbitDLLFree;
	}


	
	
}



