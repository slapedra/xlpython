/*
 2008 Sebastien Lapedra
*/


#ifndef xlp_scalar_hpp
#define xlp_scalar_hpp

#include <xlp/util/scalarvisitors.hpp>
#include <string>
#include <vector>

namespace xlp {

	//! Scalar class
  /*!
	 */
	class Scalar {
	public:

		Scalar(const ScalarDef &scalar = NullType()) : scalar_(scalar) {}
		Scalar(const Integer& i) : scalar_(i) {}
		Scalar(const Real& d) : scalar_(d) {}
    Scalar(const bool& b) : scalar_(b) {}
		Scalar(const char *s) : scalar_(std::string(s)) {}
		Scalar(const std::string& s) : scalar_(s) {}
		Scalar(const NullType& n) : scalar_(n) {}
		Scalar(const NAType& n) : scalar_(n) {}

		Scalar& operator=(ScalarDef &scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(Integer scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(Real scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(bool scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(std::string scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(NullType scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(NAType scalar) {
			scalar_ = scalar;
			return *this;
		}

		Scalar& operator=(char *scalar) {
			scalar_ = std::string(scalar);
			return *this;
		}

		//ScalarType type() const;
    bool missing() const;
		bool error() const;
		
		operator Integer() const;
		operator Real() const;
		operator bool() const;
		operator std::string() const;
		const ScalarDef& scalar() const { return scalar_; }

	private:
		ScalarDef scalar_;
	};

	/*inline ScalarType Scalar::type() const {
		return boost::apply_visitor(ScalarTypeVisitor(), scalar_);
	}*/

	inline bool Scalar::missing() const {
		return boost::apply_visitor(ScalarMissingVisitor(), scalar_);
	}

	inline bool Scalar::error() const {
		return boost::apply_visitor(ScalarErrorVisitor(), scalar_);
	}

	inline Scalar::operator Integer() const {
		return boost::apply_visitor(ScalarToIntegerVisitor(), scalar_);
  }

	inline Scalar::operator Real() const {
		return boost::apply_visitor(ScalarToRealVisitor(), scalar_);
  }

	inline Scalar::operator std::string() const {
		return boost::apply_visitor(ScalarToStringVisitor(), scalar_);
  }

	inline Scalar::operator bool() const {
		return boost::apply_visitor(ScalarToBoolVisitor(), scalar_);
  }

	inline std::ostream &operator<<(std::ostream &out, const Scalar &scalar) {
		return out << boost::apply_visitor(ScalarToStreamVisitor(), scalar.scalar());
	}

	inline std::ostream &operator<<(std::ostream &out, const ScalarType &type) {
		switch (type) {
			case ScalarInteger:
				return out << "Integer";
			case ScalarReal:
				return out << "Real";
			case ScalarBoolean:
				return out << "Boolean";
			case ScalarString:
				return out << "String";
			case ScalarNull:
				return out << "Null";
			case ScalarNA:
				return out << "N/A";
			default:
				FAIL("Unexpected type enumeration: '" << type << "'");
		}
	}


	template <class T>
	struct Range {

		Range() : rows(0), columns(0) {}
		Range(Size r, Size c) : rows(r), columns(c), data(r*c) {}
		Size size() { return rows*columns;}

		T operator()(Size i, Size j) {
			return data[i*columns+j];
		}

		Size rows;
		Size columns;
		std::vector<T> data;
	};

}

#endif

