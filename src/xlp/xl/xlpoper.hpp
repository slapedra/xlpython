/*
 2008 Sebastien Lapedra

 This file is part of XLPython, a free-software/open-source library 
 Python add-in for Excel:

 XLPython is free software: you can redistribute it and/or modify it under the
 terms of the XLPython license.  You should have received a copy of the
 license along with this program; if not, please email xlp-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifndef xlpoper_hpp
#define xlpoper_hpp

#include <xlp/xlpdefines.hpp>
#include <xlp/xl/xlutils.hpp>
#include <xlp/util/types.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/errors.hpp>
#include <vector>


namespace xlp {


	//!  Wrapper for OPER type
	/*!
		This class is useful for converting oper to scalar and scalar to oper.
	*/
	class DLL_API XlpOper {
	public:

		XlpOper(OPER *xIn, const std::string& id, bool transpose = false, 
			bool forceScalar = true, bool reduce = false);
		XlpOper(const Integer& i, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const Real& r, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const bool& b, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const std::string& s, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const char *c, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const NullType& n, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const Scalar& s, const std::string& id = "", bool transpose = false, bool dllfree = true);
		XlpOper(const boost::shared_ptr< Range<Scalar> >& sc, const std::string& id = "", 
			bool transpose = false, bool pretty = false, bool dllfree = true);

		~XlpOper();

		bool missing() const;
     
		operator Integer() const;
		operator Real() const;
		operator bool() const;
		operator std::string() const;
		operator Scalar() const;
		operator boost::shared_ptr<Range<Scalar> >() const;

		//OPER *operator->() const { return oper_; }

		OPER *get() const { 
			return oper_; 
		}

		void setToFree(bool flag) {
			flagToFree_ = flag;
		}

	private:
		OPER *oper_;
		std::string id_;
		bool forceScalar_;
		bool flagToFree_;

		boost::shared_ptr< Range<Scalar> > scalars_;
    
		void convertToScalarVector(bool transpose);
		void reduceToScalarVector(bool transpose);
		void convertToOperVector(bool transpose, bool pretty, bool dllfree = true);
		
		bool validateMulti();
		void validateRange();

		static Scalar convertNonMultiOperToScalar(const OPER *oper);
		static bool missing(const OPER *oper);
		static bool error(const OPER *oper);
		static bool scalar(const OPER *oper);

		static std::string strConv(const OPER *xString);
	};

}

#endif

