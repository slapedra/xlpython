/*
 2008 Sebastien Lapedra
*/


#ifndef xlref_hpp
#define xlref_hpp

#define NOMINMAX
#include <windows.h>
#include <xlp/xl/xldefines.hpp>
#include <xlp/xl/xlcall.h>
#include <xlp/xl/framewrk.hpp>
#include <xlp/xl/xlpoper.hpp>
#include <xlp/util/singleton.hpp>
#include <xlp/util/errors.hpp>
#include <xlp/util/scalar.hpp>
#include <xlp/util/types.hpp>
#include <sstream>
#include <stdlib.h> 
#include <xlp/xl/xloper.hpp>
#include <list>
#include <set>

namespace xlp {

	class XLRef{
	public:
		XLRef(const boost::shared_ptr<Xloper>& xloper);

		XLRef(Size row, Size col, Size rows, Size cols, long idSheet); 

		inline Size row() const { return row_; }
		inline Size rows() const { return rows_; }
		inline Size col() const { return col_; }
		inline Size cols() const { return cols_; }
		inline long idSheet() const { return idSheet_; }
		boost::shared_ptr<Xloper> oper() const { return xloper_; }

		bool inter(const boost::shared_ptr<XLRef>& xlRef) const;

	private:
		Size row_;
		Size col_;
		Size rows_;
		Size cols_;
		long idSheet_;
		boost::shared_ptr<Xloper> xloper_;

		bool in(Real x, Real y) const;
	};
}

#endif
