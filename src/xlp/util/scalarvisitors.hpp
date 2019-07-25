/*
 2008 Sebastien Lapedra
*/



#ifndef xlp_scalar_visitors_hpp
#define xlp_scalar_visitors_hpp

#include <xlp/xlpdefines.hpp>
#include <boost/variant.hpp>
#include <xlp/util/types.hpp>
#include <xlp/util/errors.hpp>
#include <boost/lexical_cast.hpp>
#include <boost/algorithm/string/trim.hpp> 

namespace xlp {

	//! Null type class
	/*!
	*/
	class NullType {
	public:
		bool operator==(NullType& n) { return true; }
		bool operator==(const NullType& n) const { return true; }
	};

	class NAType {
	public:
		bool operator==(NAType& n) { return true; }
		bool operator==(const NAType& n) const { return true; }
	};

	inline std::ostream &operator<<(std::ostream &out, const NullType &) {
		return out << "null";
	}

	inline std::ostream &operator<<(std::ostream &out, const NAType &) {
		return out << "N/A";
	}

	enum ScalarType { ScalarInteger, ScalarReal, ScalarBoolean, ScalarString, ScalarNull, ScalarNA};

	typedef boost::variant<Integer, Real, bool, std::string, NullType, NAType> ScalarDef;

	//! Convert a scalar to a Integer
	class ScalarToIntegerVisitor : public boost::static_visitor<Integer> {
	public:
		Integer operator()(const Integer &i) const { return i; }
		Integer operator()(const Real &r) const { return static_cast<Integer>(r); }
		Integer operator()(const bool &b) const { return b; }
		Integer operator()(const std::string&) const { FAIL("invalid conversion from String to Integer"); }
		Integer operator()(const NullType&) const { FAIL("invalid conversion from NullType to Integer"); }
		Integer operator()(const NAType&) const { FAIL("invalid conversion from NullType to Integer"); }
	};

	//! Convert a Scalar to a Real
	class ScalarToRealVisitor : public boost::static_visitor<Real> {
	public:
		Real operator()(const Integer &i) const { return static_cast<Real>(i); }
		Real operator()(const Real &r) const { return r; }
		Real operator()(const bool &b) const { return static_cast<Real>(b); }
		Real operator()(const std::string&) const { FAIL("invalid conversion from String to Real"); }
		Real operator()(const NullType&) const { FAIL("invalid conversion from NullType to Real"); }
		Real operator()(const NAType&) const { FAIL("invalid conversion from NullType to Real"); }
	};

	//! Convert a Scalar to a bool
	class ScalarToBoolVisitor : public boost::static_visitor<bool> {
	public:
		bool operator()(const Integer &i) const { FAIL("invalid conversion from String to Integer"); }
		bool operator()(const Real &r) const { FAIL("invalid conversion from Real to bool"); }
		bool operator()(const bool &b) const { return b; }
		bool operator()(const std::string&) const { FAIL("invalid conversion form String to bool"); }
		bool operator()(const NullType&) const { FAIL("invalid conversion form NullType to bool"); }
		bool operator()(const NAType&) const { FAIL("invalid conversion form NullType to bool"); }
	};

	//! Convert a Scalar to a string
	class ScalarToStringVisitor : public boost::static_visitor<std::string> {
	public:
		std::string operator()(const Integer &i) const {FAIL("invalid conversion from Integer to String"); }
		std::string operator()(const Real &r) const { FAIL("invalid conversion from Real to String"); }
		std::string operator()(const bool &b) const { FAIL("invalid conversion from bool to String"); }
		std::string operator()(const std::string& s) const { return s; }
		std::string operator()(const NullType&) const { return ""; }
		std::string operator()(const NAType&) const { return ""; }
	};

	//! Convert a Scalar to a stream
	class ScalarToStreamVisitor : public boost::static_visitor<std::string> {
	public:
		std::string operator()(const Integer &i) const { return boost::lexical_cast<std::string>(i); }
		std::string operator()(const Real &r) const { return boost::lexical_cast<std::string>(r); }
		std::string operator()(const bool &b) const { return boost::lexical_cast<std::string>(b); }
		std::string operator()(const std::string& s) const { return s; }
		std::string operator()(const NullType&) const { return ""; }
		std::string operator()(const NAType&) const { return ""; }
	};

	//! bool to know if the scalar is missing
	class ScalarMissingVisitor : public boost::static_visitor<bool> {
	public:
		bool operator()(const Integer&) const { return false; }
		bool operator()(const Real&) const { return false; }
		bool operator()(const bool&) const { return false; }
		bool operator()(const std::string&) const { return false; }
		bool operator()(const NullType&) const { return true; }
		bool operator()(const NAType&) const { return true; }
	};

	//! bool to know if the scalar is in error
	class ScalarErrorVisitor : public boost::static_visitor<bool> {
	public:
		bool operator()(const Integer&) const { return false; }
		bool operator()(const Real&) const { return false; }
		bool operator()(const bool&) const { return false; }
		bool operator()(const std::string&) const { return false; }
		bool operator()(const NullType&) const { return false; }
		bool operator()(const NAType&) const { return false; }
	};

	//! Return the Type of a Scalar
	class ScalarTypeVisitor : public boost::static_visitor<ScalarType> {
	public:
		ScalarType operator()(const Integer&) const { return ScalarInteger; }
		ScalarType operator()(const Real&) const { return ScalarReal; }
		ScalarType operator()(const bool&) const { return ScalarBoolean; }
		ScalarType operator()(const std::string&) const { return ScalarString; }
		ScalarType operator()(const NullType&) const { return ScalarNull; }
		ScalarType operator()(const NAType&) const { return ScalarNull; }
	};

	//! Excel id 
	class ExcelIdVisitor : public boost::static_visitor<ScalarDef> {
	public:
		ScalarDef operator()(const Integer& i) const { return i; }
		ScalarDef operator()(const Real& r) const { return r; }
		ScalarDef operator()(const bool& b) const { return b; }
		ScalarDef operator()(const std::string& s) const { return std::string(XLP_TOKENID_KEYWORD+s); }
		ScalarDef operator()(const NullType& n) const { return n; }
		ScalarDef operator()(const NAType& n) const { return n; }
	};

	//! Python id 
	class PythonIdVisitor : public boost::static_visitor<ScalarDef> {
	public:
		ScalarDef operator()(const Integer& i) const { return i; }
		ScalarDef operator()(const Real& r) const { return r; }
		ScalarDef operator()(const bool& b) const { return b; }
		ScalarDef operator()(const std::string& s) const { 
			std::string ret = s;
			boost::trim_left_if(ret, boost::is_any_of(XLP_TOKENID_KEYWORD));
			return ret;
		}
		ScalarDef operator()(const NullType& n) const { return n; }
		ScalarDef operator()(const NAType& n) const { return n; }
	};

}

#endif

