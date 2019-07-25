/*
 2008 Sebastien Lapedra
*/


#ifndef xlp_utils_hpp
#define xlp_utils_hpp

#include <boost/python.hpp>
#include <xlp/xlpdefines.hpp>
#include <boost/variant.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>

using namespace boost;

namespace xlp {


	class PyObjectScalarVisitor : public boost::static_visitor<python::object> { 
	public:
		PyObjectScalarVisitor(const python::dict& session, const python::dict& local, bool errorInfo = false)
		: session_(session), local_(local), errorInfo_(errorInfo) {
		}

		python::object operator()(const std::string& sstr) const;
		python::object operator()(const Integer& i) const;
		python::object operator()(const Real& r) const;
		python::object operator()(const bool& b) const;
		python::object operator()(const NullType& b) const;
		python::object operator()(const NAType& b) const;

	private:
		python::dict session_;
		python::dict local_;
		bool errorInfo_;
  };

}


#endif

