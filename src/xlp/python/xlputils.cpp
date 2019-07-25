/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/xlpdefines.hpp>
#include <xlp/python/xlputils.hpp>
#include <boost/lexical_cast.hpp>
#include <boost/xpressive/xpressive_static.hpp>
#include <iostream>
#include <sstream>

using namespace boost::xpressive;

namespace xlp {


	/*************************************************************************
		PyObjectScalarVisitor
	*************************************************************************/
	python::object PyObjectScalarVisitor::operator()(const std::string& sstr) const {
		static cregex regstr = XLP_TOKENID_KEYWORD >> (s1 = +_);
		cmatch what;
		if(regex_match(sstr.c_str(), what, regstr)) {
			python::str key = what[1].first;
			/*if (session_.has_key(key)) {
				return session_[key];
			} else {
				try {
					return python::eval(key, session_, local_);
				} catch(python::error_already_set&) {
					if (!errorInfo_) {
						PyErr_Print();
					} else {
						PyErr_Clear();
					}
					FAIL ("no object with id/code: " << what[1].first);
				}
				catch (...) {
				}
			}*/
			try {
				return python::eval(key, session_, local_);
			} catch(python::error_already_set&) {
				if (!errorInfo_) {
					PyErr_Print();
				} else {
					PyErr_Clear();
				}
				FAIL ("no object with id/code: " << what[1].first);
			}
		}
		return python::str(sstr.c_str());
 	}

	python::object PyObjectScalarVisitor::operator()(const Integer& i) const {
		return python::object(i);
	}

	python::object PyObjectScalarVisitor::operator()(const Real& r) const {
		return python::object(r);
	}

	python::object PyObjectScalarVisitor::operator()(const bool& b) const {
		return python::object(b);
	}

	python::object PyObjectScalarVisitor::operator()(const NullType& n) const {
		return boost::python::object(boost::python::handle<>(python::detail::none()));
	}

	python::object PyObjectScalarVisitor::operator()(const NAType& n) const {
		FAIL ("cannot convert error in python");
	}
	
}



