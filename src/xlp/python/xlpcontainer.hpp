/*
 2008 Sebastien Lapedra
*/

#ifndef xlp_container_hpp
#define xlp_container_hpp

#include <xlp/xlpdefines.hpp>
#include <boost/variant.hpp>
#include <boost/variant/apply_visitor.hpp>
#include <boost/python.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>

using namespace boost;

namespace xlp {

	extern python::str extStrCreate(const python::object& rows, 
		const python::object& columns, const python::list& datas);

	extern python::list listOfListCreate(const python::object& rows, 
		const python::object& columns, const python::list& datas);



	//!  Container ouput method
	/*!
		Generic class that define the method to create excel datas from container 
	*/
	class XLPContainerOutBindingMethod {
	public:
		virtual boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const = 0;
		
		virtual Scalar scalarOut(const python::object& obj, python::dict& session) const {
			return std::string(obj.ptr()->ob_type->tp_name);
		}
	};

	/**************************************************************************************************
		Scalar
	**************************************************************************************************/
	class XLPFloatOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "float"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};

	class XLPIntOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "int"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};

	class XLPBoolOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "bool"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
		
	};

	class XLPStrOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "str"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};


	class XLPNoneOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "NoneType"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};

	class XLPTypeOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "type"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};

	class XLPComplexOutBindingMethod : public XLPContainerOutBindingMethod {
	public:
		static std::string key() { return "complex"; }
		
		boost::shared_ptr< Range<python::object> > out(const python::object& obj, 
			python::dict& session) const;
		
		Scalar scalarOut(const python::object& obj, 
			python::dict& session) const;
	};

}

#endif

