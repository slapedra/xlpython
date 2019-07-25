/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/python/xlpcontainer.hpp>
#include <boost/tokenizer.hpp>

namespace xlp {

	/**************************************************************************************************
		Scalar
	**************************************************************************************************/
	extern python::str extStrCreate(const python::object& rrows, 
		const python::object& ccolumns, const python::list& datas) {
		Integer offset = 0;
		std::string ret = "";
		Size rows = python::extract<Size>(rrows);
		Size columns = python::extract<Size>(ccolumns);
		REQUIRE (len(datas) == rows*columns, "bad size for input datas");
		for (Size i=0;i<rows;i++) {
			for (Size j=0;j<columns;j++) {
				std::string sep = " \t";
				if (j == columns-1) {
					if (i == rows-1) {
						sep = "";
					} else {
						sep = "\n";
					}
				}
				python::object tmp = python::extract<python::object>(datas[offset]);
				std::string type = (tmp.ptr()->ob_type->tp_name);
				if (type == "NoneType") {
					ret += sep;
				} else if (type == "str"){
					ret += std::string(python::extract<std::string>(datas[offset])) + sep;
				} else {
					FAIL ("cell ("<< i << ",  " << j << ") is not of str type");
				}
				offset++;
			}
		}
		return python::str(ret);
	}

	extern python::list listOfListCreate(const python::object& rrows, 
		const python::object& ccolumns, const python::list& datas) {
		python::list obj;
		Size rows = python::extract<Size>(rrows);
		Size columns = python::extract<Size>(ccolumns);
		Integer offset = 0;
		REQUIRE (len(datas) == rows*columns, "bad size for input datas");
		for (Size i=0;i<rows;i++) {
			python::list tmp;
			for (Size j=0;j<columns;j++) {
				tmp.append(datas[offset]);
				offset++;
			}
			obj.append(tmp);
		}
		return obj;
	}

	boost::shared_ptr< Range<python::object> > XLPFloatOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		ret->data[0] = obj;
		return ret;
	}

	Scalar XLPFloatOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		return Real(python::extract<Real>(obj));
	}

	boost::shared_ptr< Range<python::object> > XLPIntOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		ret->data[0] = obj;
		return ret;
	}

	Scalar XLPIntOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		return int(python::extract<int>(obj));
	}

	boost::shared_ptr< Range<python::object> > XLPBoolOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		ret->data[0] = obj;
		return ret;
	}

	Scalar XLPBoolOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		return bool(python::extract<bool>(obj));
	}

	boost::shared_ptr< Range<python::object> > XLPStrOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		std::vector< std::vector<std::string> > datas;
		typedef boost::tokenizer<boost::char_separator<char> > tokenizer;
		boost::char_separator<char> sep("\n");
		std::string tmp = python::extract<std::string>(obj);
		tokenizer tokens(tmp, sep);
		for (tokenizer::iterator tok_iter = tokens.begin();
			tok_iter != tokens.end(); ++tok_iter) {
			std::vector<std::string> tmpv;
			tmpv.push_back(*tok_iter);
			datas.push_back( std::vector<std::string>(tmpv) );
		}
		Size columns = 0;
		boost::char_separator<char> septab("\t");
		for (Size i=0;i<datas.size();i++) {
			std::string tmp = datas[i][0];
			tokenizer tokens(tmp, septab);
			for (tokenizer::iterator tok_iter = tokens.begin();
				tok_iter != tokens.end(); ++tok_iter) {
				datas[i].push_back(*tok_iter);
			}
			if (columns < datas[i].size()-1) {
				columns = datas[i].size()-1;
			}
		}
		Integer offset = 0;
		Size rows = datas.size();
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(rows, columns));
		for (Size i=0;i<rows;i++) {
			for (Size j=0;j<datas[i].size()-1;j++) {
				ret->data[offset] = python::str(datas[i][j+1]);
				offset++;
			}
			for (Size j=datas[i].size()-1;j<columns;j++) {
				ret->data[offset] = python::str("");
				offset++;
			}
		}
		return ret;
	}

	Scalar XLPStrOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		return std::string(python::extract<std::string>(obj));
	}

	boost::shared_ptr< Range<python::object> > XLPNoneOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		ret->data[0] = obj;
		return ret;
	}

	Scalar XLPNoneOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		return NullType();
	}

	boost::shared_ptr< Range<python::object> > XLPComplexOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		std::string tmp = scalarOut(obj, session);
		ret->data[0] = python::str(tmp);;
		return ret;
	}

	Scalar XLPComplexOutBindingMethod::scalarOut(const python::object& obj, python::dict& session) const {
		std::string tmp = XLP_VAR_KEYWORD; 
		Real real = python::extract<Real>(obj.attr("real"));
		Real imag = python::extract<Real>(obj.attr("imag"));
		tmp += boost::lexical_cast<std::string>(real)+"+"+boost::lexical_cast<std::string>(imag)+"j";
		return tmp;
	}

	boost::shared_ptr< Range<python::object> > XLPTypeOutBindingMethod::out(const python::object& obj, 
		python::dict& session) const {
		boost::shared_ptr< Range<python::object> > ret(new Range<python::object>(1, 1));
		std::string tmp = XLPTypeOutBindingMethod::key();
		ret->data[0] = python::str(tmp);
		return ret;
	}
	
	Scalar XLPTypeOutBindingMethod::scalarOut(const python::object& obj, 
		python::dict& session) const {
		std::string tmp = XLPTypeOutBindingMethod::key();
		return tmp;
	}

	



}
