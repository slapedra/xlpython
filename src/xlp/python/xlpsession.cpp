/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/python/xlpsession.hpp>
#include <stdio.h>
#include <iostream>
#include <fstream>
#include <xlp/xl/xl.hpp>
#include <boost/lexical_cast.hpp>
#include <boost/format.hpp>
#include <xlp/python/xlputils.hpp>
#include <boost/multi_index_container.hpp>
#include <boost/multi_index/member.hpp>
#include <boost/multi_index/ordered_index.hpp>
#include <boost/multi_index/hashed_index.hpp>
#include <boost/xpressive/xpressive_static.hpp>
#include <boost/assign.hpp>
#include <boost/assign/std/vector.hpp>
#include <boost/python/object.hpp>
#include <boost/python/stl_iterator.hpp>
#include<boost/tokenizer.hpp>
#include <ctime>
#include <boost/crc.hpp>  
#include <boost/cstdint.hpp>


using namespace boost::assign;
using namespace boost::xpressive;

namespace xlp {


	void XLPPyerr::write(const char* text) {
		XLPSession::instance().pyerr() << text;
		XLPSession::instance().pyerr().flush();
	}

	void XLPPyout::write(const char* text) {
		XLPSession::instance().pyout() << text;
		XLPSession::instance().pyout().flush();
	}

	/******************************************************************************************
		XLPSession::ReturnHandler
	******************************************************************************************/
	python::object XLPSession::ReturnHandler::serieValue(const std::string& id, Size index) const {
		boost::shared_ptr<XLRef> ref = XLSession::instance().currentRef(caller_, id);
		Scalar value = XLSession::instance().serieValue(ref, index);
		return XLPSession::instance().scalarToPyObj(value);
	}

	python::object XLPSession::ReturnHandler::serie(const std::string& id) const {
		boost::shared_ptr<XLRef> ref = XLSession::instance().currentRef(caller_, id);
		boost::shared_ptr< Range<Scalar> > data = XLSession::instance().serie(ref);
		python::list l;
		for (Size i=0;i<data->data.size();++i) {
			l.append(XLPSession::instance().scalarToPyObj(data->data[i]));
		}
		return l;
	}

	std::string XLPSession::ReturnHandler::set(const python::object& pyobject, 
		const std::string& output, bool transpose, bool force, Real time) const {
		boost::shared_ptr<XLRef> ref = XLSession::instance().currentRef(caller_, output);
		boost::shared_ptr< Range<Scalar> > data = XLPSession::instance().get(pyobject, force);
		return XLSession::instance().fillOutput(ref, data, transpose, time);
	}

	void XLPSession::ReturnHandler::error(const std::string& msg) const {
		XLSession::instance().error(msg);
	}
	
	void XLPSession::ReturnHandler::warning(const std::string& msg) const {
		XLSession::instance().warning(msg);
	}
	
	void XLPSession::ReturnHandler::message(const std::string& msg) const {
		XLSession::instance().message(msg);
	}

	std::string XLPSession::ReturnHandler::workbookPath() const {
		return XLPSession::instance().workbookPath();
	}

	std::string XLPSession::ReturnHandler::workbookName() const {
		return XLSession::instance().workbookName();
	}

	std::string XLPSession::ReturnHandler::worksheetName() const {
		return XLSession::instance().worksheetName();
	}

	std::string XLPSession::ReturnHandler::tmpDir() const {
		return XLSession::instance().tmpDir();
	}

	void XLPSession::ReturnHandler::visible(const python::object& ssheet, 
		const python::list& pysheets, bool activate) const {
		std::vector<std::string> sheets(python::len(pysheets));
		for (Size i=0;i<sheets.size();++i) {
			sheets[i] = python::extract<std::string>(pysheets[i]);
		}
		std::string sheet;
		std::string type = (ssheet.ptr()->ob_type->tp_name);
		if (type == "str") {
			sheet = python::extract<std::string>(ssheet);
		} else {
			Integer index = python::extract<Integer>(ssheet);
			REQUIRE (index<Integer(sheets.size()), "index too big: " << index);
			REQUIRE (index>=0, "index cannot be negative");
			sheet = sheets[index];
		}
		XLSession::instance().visible(sheets, sheet, activate);
	}

	void XLPSession::ReturnHandler::name(const std::string& scalar, const std::string& name) const {
		XLSession::instance().defineName(scalar, name);
	}

	std::string XLPSession::ReturnHandler::formula(const python::object& pyobject, 
		const std::string& output, bool transpose, bool force, Real time) const {
		boost::shared_ptr<XLRef> ref = XLSession::instance().currentRef(caller_, output);
		boost::shared_ptr< Range<Scalar> > data = XLPSession::instance().get(pyobject, force);
		return XLSession::instance().formula(ref, data, transpose, time);
	}

	Real XLPSession::ReturnHandler::time() const {
		return XLSession::instance().time();
	}

	void XLPSession::ReturnHandler::save(const python::object& pyobject, const std::string& id) const {
		boost::shared_ptr< Range<Scalar> > range = XLPSession::instance().get(pyobject, true);
		XLSession::instance().saveRange(range, id);
	}

	void XLPSession::ReturnHandler::load(const std::string& id, const std::string& output, bool transpose) const {
		boost::shared_ptr< Range<Scalar> > ret = XLSession::instance().loadRange(id);
		boost::shared_ptr<XLRef> ref = XLSession::instance().currentRef(caller_, output);
		XLSession::instance().fillOutput(ref, ret, transpose);
	}

	/******************************************************************************************
		XLPSession
	******************************************************************************************/
	XLPSession::XLPSession()
	: xlpIoStreamDatas_(XLP_IOSTREAMBUFFERSIZE), 
		xlpCout_("cout", xlpIoStreamDatas_), 
		xlpCerr_("cerr", xlpIoStreamDatas_),
		xlpPyout_("pyout", xlpIoStreamDatas_), 
		xlpPyerr_("pyerr", xlpIoStreamDatas_), 
	  returnHandlerInit_(false) {
	}

	void XLPSession::initialize() { 
		// Excel init
		// Retranspose
		oldCout_ = std::cout.rdbuf();
		oldCerr_ = std::cerr.rdbuf();
		std::cout.rdbuf(xlpCout_.rdbuf());
		std::cerr.rdbuf(xlpCerr_.rdbuf());

		// Python init
		Py_OptimizeFlag = 2;
		Py_NoSiteFlag = 1;
		Py_Initialize();
		int argc = 1; 
		#if PY_VERSION_HEX >= 0x03000000
		wchar_t* argv[1] = { L"XLPython" };
		#else
		char* argv[1] = {"XLPython"};
		#endif
		PySys_SetArgv(argc, argv);

		module_ = python::import("__main__");
		session_ = python::extract<python::dict>(module_.attr("__dict__"));

		// python io retranspose
		python::object xlpPyerr = (
			python::class_<XLPPyerr>("XLPPyerr")
			.def("write", &XLPPyerr::write)) ();

		python::object xlpPyout = (
			python::class_<XLPPyout>("XLPPyout")
			.def("write", &XLPPyout::write)) ();

		python::object sys = python::import("sys");
		sys.attr("stdout") = xlpPyout;
		sys.attr("stderr") = xlpPyerr;
		
		// out
		containerOutBinds_.insert(std::make_pair(XLPFloatOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPFloatOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPBoolOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPBoolOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPStrOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPStrOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPIntOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPIntOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPNoneOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPNoneOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPTypeOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPTypeOutBindingMethod())));

		containerOutBinds_.insert(std::make_pair(XLPComplexOutBindingMethod::key(), 
			boost::shared_ptr<XLPContainerOutBindingMethod>(new XLPComplexOutBindingMethod())));
  }

	XLPSession::~XLPSession() {
		//Py_Finalize();
		std::cout.rdbuf(oldCout_);
		std::cerr.rdbuf(oldCerr_);
	}

	void XLPSession::validateObjectId(const std::string& id) const {
		static sregex regstr = alpha >> *_w;
		smatch what;
		if(regex_match(id, what, regstr)) {
		} else {
			FAIL ("invalid id: " << id);
		}
	}

	std::string XLPSession::createObjectId(const python::object& obj, const std::string& id) const {
		boost::shared_ptr<XLCurrentPosition> pos = XLSession::instance().currentPosition();
		std::string key;
		if (id.empty()) {
			key = XLP_OBJ_KEYWORD + pos->idPosition()+ XLP_TOKENID_KEYWORD + pos->idSheet();
		} else {	
			key = id + XLP_TOKENID_KEYWORD + pos->idSheet();
			validateObjectId(key);
		}
		return key;
	}

	std::string XLPSession::createExcelId(const std::string& id) const {	
		std::string key = XLP_TOKENID_KEYWORD + id;
		return key;
	}

	std::string XLPSession::getType(const python::object& obj) const {
		python::object type = obj.attr("__class__" );
		std::string ret = python::extract<std::string>(python::str(type));
		typedef boost::tokenizer<boost::char_separator<char> > tokenizer;
		boost::char_separator<char> sep("'");
		tokenizer tokens(ret, sep);
		tokenizer::iterator tok_iter = tokens.begin();
		++tok_iter;
		return *tok_iter;
	}

	std::string XLPSession::pythonPath() {
		char *path = getenv("PYTHONPATH");
		if (path != NULL) {
			return std::string(path);
		} else {
			FAIL ("PYTHONPATH not defined");
		}
	}

	std::string XLPSession::pythonVersion() {
		try {
			python::object sys = python::import("sys");
			std::string version = python::extract<std::string>(sys.attr("version"));	
			return version;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::string XLPSession::workbookPath() {
		return XLSession::instance().workbookPath();
	}

	std::string XLPSession::workbookName() {
		return XLSession::instance().workbookName();
	}

	std::string XLPSession::worksheetName(bool little) {
		return XLSession::instance().worksheetName(little);
	}
	
	std::string XLPSession::type(const Scalar& scalar) const {
		try {
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, session_), scalar.scalar());
			return getType(obj);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::string XLPSession::handleException() const {
		std::string msg;
		if (PyErr_ExceptionMatches(PyExc_AssertionError)) 
			msg = "AssertionError";
		else if (PyErr_ExceptionMatches(PyExc_AttributeError)) 
			msg = "AttributeError ";
    else if (PyErr_ExceptionMatches(PyExc_EOFError)) 
			msg = "EOFError ";
    else if (PyErr_ExceptionMatches(PyExc_FloatingPointError)) 
			msg = "FloatingPointError ";
    else if (PyErr_ExceptionMatches(PyExc_IOError)) 
			msg = "IOError";
		else if (PyErr_ExceptionMatches(PyExc_ImportError)) 
			msg = "ImportError";
		else if (PyErr_ExceptionMatches(PyExc_IndexError)) 
			msg = "IndexError";
    else if (PyErr_ExceptionMatches(PyExc_KeyError)) 
			msg = "KeyError";
    else if (PyErr_ExceptionMatches(PyExc_KeyboardInterrupt)) 
			msg = "KeyboardInterrupt "; 
	  else if (PyErr_ExceptionMatches(PyExc_MemoryError)) 
			msg = "MemoryError";
    else if (PyErr_ExceptionMatches(PyExc_NameError)) 
			msg = "NameError";
    else if (PyErr_ExceptionMatches(PyExc_NotImplementedError)) 
			msg = "NotImplementedError";
    else if (PyErr_ExceptionMatches(PyExc_OSError)) 
			msg = "OSError";
		else if (PyErr_ExceptionMatches(PyExc_OverflowError)) 
			msg = "OverflowError";
    else if (PyErr_ExceptionMatches(PyExc_ReferenceError)) 
			msg = "ReferenceError";
    else if (PyErr_ExceptionMatches(PyExc_RuntimeError)) 
			msg = "RuntimeError";
    else if (PyErr_ExceptionMatches(PyExc_SyntaxError)) 
			msg = "SyntaxError";
    else if (PyErr_ExceptionMatches(PyExc_SystemError)) 
			msg = "SystemError";
		else if (PyErr_ExceptionMatches(PyExc_SystemExit)) 
			msg = "SystemExit";
		else if (PyErr_ExceptionMatches(PyExc_TypeError)) 
			msg = "TypeError";
		else if (PyErr_ExceptionMatches(PyExc_ValueError)) 
			msg = "ValueError";
		else if (PyErr_ExceptionMatches(PyExc_WindowsError)) 
			msg = "WindowsError";
		else if (PyErr_ExceptionMatches(PyExc_ZeroDivisionError)) 
			msg = "ZeroDivisionError";
		PyErr_Print();
		PyErr_Clear();
		return msg;
	}

	XLPSession::CallablePyObjectInfos XLPSession::getCallableObjectInfos(const python::object& obj) const {
		try {
			CallablePyObjectInfos ret;
			ret.minArgs = 0;
			ret.callable = false;
			ret.variadic = false;
			ret.instanceMethod = false;
			try {
				python::object tmp1 = obj.attr("__call__");
				ret.callable = true;
			} catch (python::error_already_set&) {
				PyErr_Clear();
			}
			try {
				python::object tmp2 = obj.attr("func_code");
				python::object tmp3 = tmp2.attr("co_argcount");
				python::object tmp4 = tmp2.attr("co_flags");
				python::object tmp5 = obj.attr("__name__");
				ret.minArgs = python::extract<Integer>(tmp3);
				Integer flags = python::extract<Integer>(tmp4);
				if (flags & 4) {
					ret.variadic = true;
				}
				ret.name = python::extract<std::string>(tmp5);
			} catch (python::error_already_set&) {
				PyErr_Clear();
			}
			try {
				python::object tmp6 = obj.attr("im_self");
				std::string type = getType(tmp6);
				if (type != "NoneType") {
					ret.minArgs--;
					ret.instanceMethod = true;
				}
			} catch (python::error_already_set&) {
				PyErr_Clear();
			}
			return ret;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}


	boost::shared_ptr<python::dict> XLPSession::localDict(
		const boost::shared_ptr< Range<Scalar> >& args,
		const boost::shared_ptr< Range<Scalar> >& values, 
  	const boost::shared_ptr< Range<Scalar> >& index) const {
		if (index->size() !=0) {
			REQUIRE (values->size() == index->size(), "values and index size not equal");
		}
		try {
			boost::shared_ptr<python::dict> local(new python::dict());
			(*local)["_extstr"] = extStrCreate;
			(*local)["_listlist"] = listOfListCreate;
			python::list argsList;
			for (Size i=0;i<args->size();i++) {
				std::string key;
				key = XLP_VAR_KEYWORD+boost::lexical_cast<std::string>(i+1);
				python::object tmp = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), 
					 args->data[i].scalar());
				(*local)[key] = tmp;
				argsList.append(tmp);
			}
			(*local)["_args"] = argsList;
			if (values->size() == 0) {
				(*local)["_rows"] = args->rows;
				(*local)["_cols"] = args->columns;
				(*local)["_size"] = args->size();
			} else {
				(*local)["_rows"] = values->rows;
				(*local)["_cols"] = values->columns;
				(*local)["_size"] = values->size();
			}
			python::list valuesList;
			for (Size i=0;i<values->size();i++) {
				python::object tmp = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), 
					 values->data[i].scalar());
				valuesList.append(tmp);
			}
			(*local)["_values"] = valuesList;

			python::list indexList;
			for (Size i=0;i<values->size();i++) {
				if (index->size() == 0) {
					indexList.append(i);
				} else {
					python::object tmp = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), 
						index->data[i].scalar());
					indexList.append(tmp);
				}
			}
			(*local)["_index"] = indexList;
			return local;
		} catch(python::error_already_set&) {
		 XLPSESSIONFAIL ("python error");
		}
	}

	void XLPSession::insert(const python::object& obj, const python::dict& local) const {
		
		try {
			python::list valuesList = python::extract<python::list>(local["_values"]);
			python::list indexList = python::extract<python::list>(local["_index"]);
			/*try {;;
		  	python::object ignored = ret.attr("__len__"); 
	  		} catch(python::error_already_set&) {
		  XLPSESSIONFAIL (type << " does not support len attribute");
	  	  }*/
			//REQUIRE (len(obj) == len(valuesList), "mismath between object and values sizes");
			python::stl_input_iterator<python::object> iterValues(valuesList), endValues;
			python::stl_input_iterator<python::object> iterIndex(indexList), endIndex;
			while (iterValues != endValues) {
				 python::object ignored = obj.attr("__setitem__")(*iterIndex, *iterValues);
				 iterIndex++;
				 iterValues++;
		  }
	  } catch(python::error_already_set&) {
			XLPSESSIONFAIL (getType(obj) << " does not support input iterator");
	  }
  }

	std::string XLPSession::set(const Scalar& scalar, 
		const boost::shared_ptr< Range<Scalar> >& args,
		const boost::shared_ptr< Range<Scalar> >& values,
		const boost::shared_ptr< Range<Scalar> >& index,
	 	const std::string& id) const {
		try {
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
			if (values->size() != 0) {
				insert(obj, *local);
			}
			std::string key = createObjectId(obj, id);
			//session_[key] = obj;
			//session_[key] = XLSession::instance().currentPosition()->workbook();
			insertObject(obj, key, XLSession::instance().currentPosition()->workbook());
			return createExcelId(key);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::get(const Scalar& scalar, 
		const boost::shared_ptr< Range<Scalar> >& args, bool force) const {
		try {
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
		  return get(obj, force);
		} catch(python::error_already_set&) {
		  XLPSESSIONFAIL ("python error");
		}
	}


	python::object XLPSession::applyScalar(const Scalar& scalar,
		const boost::shared_ptr< Range<Scalar> >& args) const {
		#ifdef XLPYTHON_INTROSPECTION
		try {
			python::object res;
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
			CallablePyObjectInfos infos = getCallableObjectInfos(obj);
			(*local)["_obj"] = obj;
			if (infos.callable) {
				if (infos.variadic) {
					REQUIRE (Size(infos.minArgs) <= args->size(), "expect at least " << infos.minArgs << "arguments");
				} else {
					REQUIRE (infos.minArgs == args->size(), "expect " << infos.minArgs << " arguments");
				}
				std::string code = "_obj(*_args)";				
				res = python::eval(code.c_str(), session_, *local);
			} else {
				REQUIRE (values->size() == 0, "not callable object");
				res = obj;
			}
			return res;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
		#else
		try {
			python::object res;
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
			CallablePyObjectInfos infos = getCallableObjectInfos(obj);
			(*local)["_obj"] = obj;
			if (infos.callable) {
				std::string code = "_obj(*_args)";				
				res = python::eval(code.c_str(), session_, *local);
			} else {
				REQUIRE (values->size() == 0, "not callable object");
				res = obj;
			}
			return res;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
		#endif
	}

	std::string XLPSession::apply(const Scalar& scalar,
		const boost::shared_ptr< Range<Scalar> >& args,
	 	const std::string& id) const {
	   	try {
		   python::object res = applyScalar(scalar, args);
		   std::string key = createObjectId(res, id);
		   //session_[key] = res;
			 	insertObject(res, key, XLSession::instance().currentPosition()->workbook());
		   return createExcelId(key);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
  }

	boost::shared_ptr< Range<Scalar> > XLPSession::getApply(const Scalar& scalar,
		const boost::shared_ptr< Range<Scalar> >& args, bool force) const {
		try {
			python::object res = applyScalar(scalar, args);
			return get(res, force);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	python::object XLPSession::attrScalar(const Scalar& scalar, const std::string& attr,
		const boost::shared_ptr< Range<Scalar> >& args) const {
		#ifdef XLPYTHON_INTROSPECTION
		try {
			python::object res;
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object tmp = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
			python::object obj;
			try {
				obj = tmp.attr(attr.c_str());
			} catch (python::error_already_set&) {
				PyErr_Clear();
				FAIL (attr.c_str() << " is not an attribute");
			}
			CallablePyObjectInfos infos = getCallableObjectInfos(obj);
			(*local)["_obj"] = obj;
			if (infos.callable) {
				if (infos.variadic) {
					REQUIRE (Size(infos.minArgs) <= args->size(), "expect at least " << infos.minArgs << "arguments");
				} else {
					REQUIRE (infos.minArgs == args->size(), "expect " << infos.minArgs << " arguments");
				}
				std::string code = "_obj(*_args)";				
				res = python::eval(code.c_str(), session_, *local);
			} else {
				REQUIRE (values->size() == 0, "not callable object");
				res = obj;
			}
			return res;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
		#else
		try {
			python::object res;
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			python::object tmp = boost::apply_visitor(PyObjectScalarVisitor(session_, *local), scalar.scalar());
			python::object obj;
			try {
				obj = tmp.attr(attr.c_str());
			} catch (python::error_already_set&) {
				PyErr_Clear();
				FAIL (attr.c_str() << " is not an attribute");
			}
			CallablePyObjectInfos infos = getCallableObjectInfos(obj);
			(*local)["_obj"] = obj;
			if (infos.callable) {
				std::string code = "_obj(*_args)";				
				res = python::eval(code.c_str(), session_, *local);
			} else {
				REQUIRE (values->size() == 0, "not callable object");
				res = obj;
			}
			return res;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
		#endif
	}

	std::string XLPSession::attr(const Scalar& scalar, const std::string& attr,
		const boost::shared_ptr< Range<Scalar> >& args,
		const std::string& id) const {
		try {
			python::object res = attrScalar(scalar, attr, args);
			std::string key = createObjectId(res, id);
			insertObject(res, key, XLSession::instance().currentPosition()->workbook());
			return createExcelId(key);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}
	
	boost::shared_ptr< Range<Scalar> > XLPSession::getAttr(const Scalar& scalar, const std::string& attr,
		const boost::shared_ptr< Range<Scalar> >& args, bool force) const {
		try {
			python::object res = attrScalar(scalar, attr, args);
			return get(res, force);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	// conversion of python object
	Scalar XLPSession::convertPyObjectToScalar(const python::object& obj, bool force) const {
		Scalar ret;
		python::object tmpObj = obj;
		std::string type = getType(tmpObj);
		OutMap::const_iterator iter = containerOutBinds_.find(type);
		if (iter == containerOutBinds_.end() && (force)) {
			tmpObj = forceConversion(obj);
			type = getType(tmpObj);
			iter = containerOutBinds_.find(type);
		}
		if (iter != containerOutBinds_.end()) {
			ret = iter->second->scalarOut(tmpObj, session_);
		} else {
			std::string tmp = searchName(tmpObj);
			if (tmp.empty()) {
				tmp = createObjectId(obj, "");
			}
			insertObject(obj, tmp, XLSession::instance().currentPosition()->workbook());
			ret = createExcelId(tmp);
		}
		return ret;
	}

	python::object XLPSession::forceConversion(const python::object& obj) const {
		try {
			python::object tmp = obj.attr("__call__")();
			return tmp;
		} catch(python::error_already_set&) {
			PyErr_Clear();
		}
		/*try {
			python::object tmp = obj.attr("__float__")();
			return tmp;
		} catch(python::error_already_set&) {
			PyErr_Clear();
		}
		try {
			python::object tmp = obj.attr("__int__")();
			return tmp;
		} catch(python::error_already_set&) {
			PyErr_Clear();
		}
		try {
			python::object tmp = obj.attr("__str__")();
			return tmp;
		} catch(python::error_already_set&) {
			PyErr_Clear();
		}*/
		FAIL ("impossible conversion");
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::get(const python::object& obj, bool force) const {
		try {
			boost::shared_ptr<XLCurrentPosition> pos = XLSession::instance().currentPosition();
			boost::shared_ptr< Range<Scalar> > ret;
			python::object tmpObj = obj;
			std::string type = getType(tmpObj);
			OutMap::const_iterator iter = containerOutBinds_.find(type);
			if (iter != containerOutBinds_.end()) {
				boost::shared_ptr< Range<python::object> > robjs = iter->second->out(tmpObj, session_);
				ret.reset(new Range<Scalar>(robjs->rows, robjs->columns));
				Size index = 0;
				for (Size i=0;i<ret->rows;i++) {
					for (Size j=0;j<ret->columns;j++) {
						ret->data[index] = convertPyObjectToScalar(robjs->data[index]);
						pos->addCol();
						++index;
					}
					pos->addRow();
				}
				return ret;
			}
			try {
				try {
					python::object ignored = obj.attr("__getitem__"); 
				} catch(python::error_already_set&) {
					PyErr_Clear();
					FAIL ("object does not support getitem");
				}
				Size size = len(obj);
				Size rows = 1;
				Size columns = size;
				try {
					Size size1 = python::extract<Size>(obj.attr("size1")());
					Size size2 = python::extract<Size>(obj.attr("size2")());
					rows  = size1;
					columns  = size2;
				} catch(python::error_already_set&) {
					PyErr_Clear();
				}
				REQUIRE (size == rows*columns, "invalid size");
				ret.reset(new Range<Scalar>(rows, columns));
				Size index = 0;
				python::object tmp;
				for (Size i=0;i<ret->rows;i++) {
					for (Size j=0;j<ret->columns;j++) {
						tmp = obj.attr("__getitem__")(index);
						ret->data[index] = convertPyObjectToScalar(tmp, force);
						pos->addCol();
						++index;
					}
					pos->addRow();
				}
				return ret;
			} catch(Error&) {
				try {
					tmpObj = forceConversion(obj);
					type = getType(tmpObj);
					iter = containerOutBinds_.find(type);
					if (iter != containerOutBinds_.end()) {
						boost::shared_ptr< Range<python::object> > robjs = iter->second->out(tmpObj, session_);
						ret.reset(new Range<Scalar>(robjs->rows, robjs->columns));
						Size index = 0;
						for (Size i=0;i<ret->rows;i++) {
							for (Size j=0;j<ret->columns;j++) {
								ret->data[index] = convertPyObjectToScalar(robjs->data[index]);
								pos->addCol();
								++index;
							}
							pos->addRow();
						}
					}
					return ret;
				} catch(Error&) {
					FAIL ("imposible to get object");
				}
			}
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::string XLPSession::import(const std::string& module, const std::string& id) const {
		try {
			python::str mod(module);
			python::object obj = python::import(mod);
			std::string key;
			if (id.empty()) {
				key = createObjectId(obj, module);
			} else {
				key = createObjectId(obj, id);
			}
			insertObject(obj, key, XLSession::instance().currentPosition()->workbook());
			return createExcelId(key);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	void XLPSession::fromImport(const std::string& module) const {
		try {
			std::string code = "from "+module+" import * ";
			python::exec(code.c_str(), session_, session_);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	bool XLPSession::deleteObject(const std::string& id)  const{
		static std::vector<std::string> frozenObjectIds;
		static bool init = false;
		if (!init) {
			frozenObjectIds +=
			std::string("__builtins__"),
			std::string("__name__"),
			std::string("__doc__"),
			std::string("__package__");
			init = true;
		}
		for (Size i=0;i<frozenObjectIds.size();i++) {
			if (id == frozenObjectIds[i]) 
				return false;
		}
		try {
			del(session_[id]);
			//PyObjectBimap::left_const_iterator iter = workbookObjects_.left.find(id);
			//if (iter != workbookObjects_.left.end()) {
			workbookObjects_.left.erase(id);
			//}
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("error deleting " << id);
		}
		return true;
	}

	void XLPSession::deleteAllObjects() const {
		python::object keys = session_.keys();
		Size size = len(keys);
		std::string id;
		for (Size i=0;i<size;i++) {
			try {
				id = std::string(python::extract<std::string>(keys[i]));
				deleteObject(id);
			} catch(python::error_already_set&) {
				XLPSESSIONFAIL ("python error");
			}
		} 
	}

	void XLPSession::onWorkbookClose(const std::string& name) const {
		PyObjectBimap::right_const_iterator up = workbookObjects_.right.upper_bound(name);
		PyObjectBimap::right_const_iterator low = workbookObjects_.right.lower_bound(name);
		PyObjectBimap::right_const_iterator iter = low;
		while (iter != up) {
			if (iter->first == name) {
				std::string id = iter->second;
				++iter;
				deleteObject(id);
			} else {
				++iter;
			}
		}
		CodeMap::iterator iter2 = codeMap_.begin();
		while (iter2 != codeMap_.end()) {
			if (iter2->workbook == name) {
				iter2 = codeMap_.erase(iter2);
			} else {
				++iter2;
			}
		}
	}

	void XLPSession::deleteObject(const Scalar& scalar) const {
		try {
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, session_), scalar.scalar());
			std::string id = scalar;
			bool res = deleteObject(id);
			if (!res) {
				FAIL (id << " is a system object");
			}
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

		boost::shared_ptr< Range<Scalar> > XLPSession::attrInfos(const Scalar& scalar, const std::string& attr) const {
		try {
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, session_), scalar.scalar());
			python::object tmp;
			try {
				tmp = obj.attr(attr.c_str());
			} catch (python::error_already_set&){
				PyErr_Clear();
				FAIL (attr.c_str() << " is not an attribute");
			}
			CallablePyObjectInfos infos = getCallableObjectInfos(tmp);
			boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(1, 4));
			ret->data[0] = getType(tmp);
			ret->data[1] = infos.minArgs;
			ret->data[2] = infos.callable;
			ret->data[3] = infos.variadic;
			return ret;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}	

	boost::shared_ptr< Range<Scalar> > XLPSession::listAllAttr(const Scalar& scalar) const {
		try {
			python::dict local;
			python::object obj = boost::apply_visitor(PyObjectScalarVisitor(session_, session_), scalar.scalar());
			local["_1"] = obj;
			python::object res = python::eval("dir(_1)", session_, local);
			python::list tmp = python::extract<python::list>(res);
			boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(len(tmp), 5));
			Size offset = 0;
			for (Size i=0;i<Size(len(tmp));i++) {
				std::string attr = std::string(python::extract<std::string>(tmp[i]));
				python::object tmp2 = obj.attr(attr.c_str());
				CallablePyObjectInfos infos = getCallableObjectInfos(tmp2);
				ret->data[offset++] = attr;
				ret->data[offset++] = getType(tmp2);
				ret->data[offset++] = infos.minArgs;
				ret->data[offset++] = infos.callable;
				ret->data[offset++] = infos.variadic;
			}
			return ret;
		} catch (python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::string XLPSession::loadFile(const std::string& filename, const std::string& id) const {
		std::string tmp;
		std::ifstream file;
		std::stringstream sstr;
		try {
			file.open(filename.c_str());
			sstr << file.rdbuf();
			tmp = sstr.str();
			file.close();
		} catch(...) {
			file.close();
			FAIL ("error opening file: " << filename);
		}	
		try {
			python::str script(tmp);
			std::string key = createObjectId(script, id);
			//session_[key] = script;
			insertObject(script, key, XLSession::instance().currentPosition()->workbook());
			return createExcelId(key);
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::exec(
		const boost::shared_ptr< Range<Scalar> >& scalarCode, 
		const boost::shared_ptr< Range<Scalar> >& args, 
		const boost::shared_ptr< Range<Scalar> >& ids) const {
		REQUIRE(scalarCode->data.size() > 0, "need scalar code");
		try {
			boost::shared_ptr<XLCurrentPosition> pos = XLSession::instance().currentPosition();
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			if (!returnHandlerInit_) {
				(*local)["XLPReturnHandler"] = python::class_<ReturnHandler>("XLPReturnHandler")
				.def(python::self += python::object())
				.def("set", &ReturnHandler::set)
				.def("set", &ReturnHandler::simplifiedSet1)
				.def("set", &ReturnHandler::simplifiedSet2)
				.def("set", &ReturnHandler::simplifiedSet3)
				.def("error", &ReturnHandler::error)
				.def("warning", &ReturnHandler::warning)
				.def("message", &ReturnHandler::message)
				.def("name", &ReturnHandler::name)
				.def("formula", &ReturnHandler::set)
				.def("formula", &ReturnHandler::simplifiedFormula1)
				.def("formula", &ReturnHandler::simplifiedFormula2)
				.def("formula", &ReturnHandler::simplifiedFormula3)
				.def("workbookPath", &ReturnHandler::workbookPath)
				.def("workbookName", &ReturnHandler::workbookName)
				.def("worksheetName", &ReturnHandler::worksheetName)
				.def("tmpDir", &ReturnHandler::tmpDir)
				.def("visible", &ReturnHandler::visible)
				.def("visible", &ReturnHandler::simplifiedVisible1)
				.def("visible", &ReturnHandler::simplifiedVisible2)
				.def("time", &ReturnHandler::time)
				.def("save", &ReturnHandler::save)
				.def("load", &ReturnHandler::load)
				.def("load", &ReturnHandler::simplifiedLoad)
				.def("serieValue", &ReturnHandler::serieValue)
				.def("serie", &ReturnHandler::serie)
				;
				returnHandlerInit_ = true;
			} 
			// create return handler
			ReturnHandler returnHandler;
			(*local)[XLP_VAR_KEYWORD] = python::ptr(&returnHandler);
			std::string id = scalarCode->data[0];
			CodeMap::iterator iter = codeMap_.find(boost::make_tuple(id, pos->workbook()));
			if (iter != codeMap_.end()) {
				PyObject* result = PyEval_EvalCode((PyCodeObject *)(iter->code.ptr()), 
					session_.ptr(), (*local).ptr());
				if (!result) python::throw_error_already_set();
			} else {
				std::ostringstream sstr;
				for (Size i=0;i<scalarCode->size();++i) {
					std::string tmp = scalarCode->data[i];
					sstr << tmp << std::endl;
				}
				python::str script(sstr.str());
				python::object ignored = python::exec(script, session_, *local);
			}
			// retrieve data
			python::list ret = returnHandler.objects();
			boost::shared_ptr< Range<Scalar> > objs( new Range<Scalar>(1, len(ret)));
			for (Size i=0;i<Size(len(ret));i++) {
				std::string key;
				if (i < ids->data.size()) {
					key = createObjectId(ret[i], ids->data[i]);
				} else {
					key = createObjectId(ret[i], "");
				}
				pos->addCol();
				//session_[key] = ret[i];
				insertObject(ret[i], key, XLSession::instance().currentPosition()->workbook());
				objs->data[i] = createExcelId(key);
			}
			return objs;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	void XLPSession::compile(
		const boost::shared_ptr< Range<Scalar> >& scalarCode, const std::string& id) const {
		REQUIRE (!id.empty(), "empty id not allowed");
		try{
			std::ostringstream sstr;
			std::vector<std::string> objIds;
			boost::shared_ptr< Range<Scalar> > args(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > values(new Range<Scalar>(0,0));
			boost::shared_ptr< Range<Scalar> > index(new Range<Scalar>(0,0));
			boost::shared_ptr<python::dict> local = localDict(args, values, index);
			for (Size i=0;i<scalarCode->size();++i) {
				std::string tmp = scalarCode->data[i];
				sstr << tmp << std::endl;
				/*std::string tmp;
				python::object obj;
				bool noCode = true;
				try {
					obj = boost::apply_visitor(PyObjectScalarVisitor(session_, *local, true), 
						scalarCode->data[i].scalar());
					std::string type = getType(obj);
					if (type == "str") {
						noCode = false;
					}
				} catch(Error&) {		
				}
				if (noCode)  {
					tmp = scalarCode->data[i];
				} else {
					tmp = python::extract<std::string>(obj);
				}
				sstr << tmp << std::endl;*/
			}
			python::str str(sstr.str());
			char *s = python::extract<char *>(str);
			PyObject* result = Py_CompileString(s, "", Py_file_input);
			if (!result) python::throw_error_already_set();
			python::object res = python::object(python::handle<>(result));
			std::string wb = XLSession::instance().currentPosition()->workbook();
			CodeMap::iterator iter = codeMap_.find(boost::make_tuple(id, wb));
			Code code(res, id, wb);
			if (iter == codeMap_.end()) {
				codeMap_.insert(code);
			} else {
				codeMap_.replace(iter, code);
			}
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::listAllObjects() const {
		try {
			python::list keys = session_.keys();
			python::list values = session_.values();
			boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(len(keys),2));
			Size offset = 0;
			for (Size i=0;i<Size(len(keys));i++) {
				ret->data[offset++] = python::extract<std::string>(keys[i]);
				ret->data[offset++] = getType(values[i]);
			}
			return ret;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::string XLPSession::searchName(const python::object& obj) const {
		try {
			python::object values = session_.values();
			python::object keys = session_.keys();
			Size size = len(session_);
			for (Size i=0;i<size;i++) {
				python::object tmp = python::extract<python::object>(values[i]);
				if (obj.ptr() == tmp.ptr()) {
					return std::string(python::extract<std::string>(keys[i]));
				}	
			}
			return "";
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	std::vector<std::string> XLPSession::getSysPath() const {
		python::object sys = python::import("sys");
		python::list res = python::extract<python::list>(sys.attr("path"));
		std::vector<std::string> ret;
		for (Size i=0;i<Size(len(res));++i) {
			ret.push_back(python::extract<std::string>(res[i]));
		}
		return ret;
	}

	void XLPSession::insertPaths(const std::vector<std::string>& paths) const {
		try {
			python::object sys = python::import("sys");
			//session_["sys"] = sys; 
			bool inFlag = false;
			std::vector<std::string> sysPaths = getSysPath();
			for (Size j=0;j<paths.size();j++) {
				for (Size i=0;i<sysPaths.size();++i) {
					if (sysPaths[i] == paths[j]) {
						inFlag = true;
						break;
					}
				}
				if (!inFlag) {
					sys.attr("path").attr("insert")(0, paths[j]);	
				}
			}
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::sysPath() const {
		try {
			bool inFlag = false;
			std::vector<std::string> paths = getSysPath();
			boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(1, paths.size()));
			for (Size i=0;i<paths.size();++i) {
				ret->data[i] = paths[i];
			}	
			return ret;
		} catch(python::error_already_set&) {
			XLPSESSIONFAIL ("python error");
		}
	}

	void XLPSession::insertObject(const python::object& obj, 
		const std::string& id, const std::string& workbook) const {
		session_[id] = obj;
		workbookObjects_.insert(PyObjectBimapElement(id, workbook));
	}

	std::string XLPSession::workbook(const std::string& id) const {
		PyObjectBimap::left_const_iterator iter = workbookObjects_.left.find(id);
		REQUIRE (iter != workbookObjects_.left.end(), id << " does not belong to a workbook");
		return iter->second;
	}

	boost::shared_ptr< Range<Scalar> > XLPSession::excelId(
		const boost::shared_ptr< Range<Scalar> >& range) const {
		boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(range->rows, range->columns));
		for (Size i=0;i<range->size();++i) {
			ret->data[0] = boost::apply_visitor(ExcelIdVisitor(), range->data[i].scalar());
		}
		return ret;
	}
	
	boost::shared_ptr< Range<Scalar> > XLPSession::pythonId(
		const boost::shared_ptr< Range<Scalar> >& range) const {
		boost::shared_ptr< Range<Scalar> > ret(new Range<Scalar>(range->rows, range->columns));
		for (Size i=0;i<range->size();++i) {
			ret->data[0] = boost::apply_visitor(PythonIdVisitor(), range->data[i].scalar());
		}
		return ret;
	}

	inline python::object XLPSession::scalarToPyObj(const Scalar& scalar) const {
		return boost::apply_visitor(PyObjectScalarVisitor(session_, session_), scalar.scalar());
	}

}



