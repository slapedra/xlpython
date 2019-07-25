/*
 2008 Sebastien Lapedra
*/

#include <xlp/config.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlpoper.hpp>
#include <sstream>
#include <boost/tokenizer.hpp>
#include <boost/xpressive/xpressive_static.hpp>
#include <xlp/version.hpp>
#include <boost/algorithm/string.hpp> 
#include <boost/algorithm/string/trim.hpp> 

#include <iomanip>
#include <iostream>
#include <fstream>
#include <boost/archive/tmpdir.hpp>
#include <boost/archive/text_iarchive.hpp>
#include <boost/archive/text_oarchive.hpp>
#include <boost/serialization/base_object.hpp>
#include <boost/serialization/utility.hpp>
#include <boost/serialization/list.hpp>
#include <boost/serialization/assume_abstract.hpp>
#include <boost/serialization/vector.hpp>
#include <boost/serialization/variant.hpp>
#include <boost/timer.hpp>


namespace xlp {

	class XLRefOper : public Xloper {
	public:
		XLRefOper()
		: Xloper(false) {
		}

    ~XLRefOper() {
			delete xloper_.val.mref.lpmref;
    }
     
	};

	XLRef::XLRef(const boost::shared_ptr<Xloper>& xloper)
	: xloper_(xloper) {
		static XLOPER xlInt;
		xlInt.xltype = xltypeInt;
		xlInt.val.w = xltypeRef;
		OPER* xlRef = new OPER;
		Excel4(xlCoerce, xlRef, 2, &(*xloper_), &xlInt);
		REQUIRE (xlRef->val.mref.lpmref->count == 1, "multiple reference not supported");
		idSheet_ = xlRef->val.mref.idSheet;
		row_ = xlRef->val.mref.lpmref->reftbl[0].rwFirst;
		rows_ = xlRef->val.mref.lpmref->reftbl[0].rwLast - row_ + 1;
		col_ = xlRef->val.mref.lpmref->reftbl[0].colFirst;
		cols_ = xlRef->val.mref.lpmref->reftbl[0].colLast - col_ + 1;
		Excel(xlFree, 0, 1, xlRef);
	}

	XLRef::XLRef(Size row, Size col, Size rows, Size cols, long idSheet)
	: xloper_(new XLRefOper()) {
		idSheet_ = idSheet;
		row_ = row;
		rows_ = rows;
		col_ = col;
		cols_ = cols;
		(&(*(xloper_.get())))->val.mref.lpmref = new XLMREF;
		(&(*(xloper_.get())))->xltype = xltypeRef;
		(&(*(xloper_.get())))->val.mref.lpmref->count = 1;
		(&(*(xloper_.get())))->val.mref.idSheet = idSheet;
		(&(*(xloper_.get())))->val.mref.lpmref->reftbl[0].rwFirst = row;
		(&(*(xloper_.get())))->val.mref.lpmref->reftbl[0].rwLast = row + rows - 1;
		(&(*(xloper_.get())))->val.mref.lpmref->reftbl[0].colFirst = col;
		(&(*(xloper_.get())))->val.mref.lpmref->reftbl[0].colLast = col + cols - 1;
	}

	bool XLRef::inter(const boost::shared_ptr<XLRef>& xlRef) const {
		if (idSheet_ != xlRef->idSheet()) {
			return false;
		} else {
			if (in(xlRef->col(), xlRef->row())) {
				return true; 
			} else if (in(xlRef->col()+xlRef->cols()-1, xlRef->row())) {
				return true;
			} else if (in(xlRef->col()+xlRef->cols()-1, xlRef->row()+xlRef->rows()-1)) {
				return true;
			} else if (in(xlRef->col(), xlRef->row()+xlRef->rows()-1)) {
				return true;
			} else {
				if ( (col_ >= xlRef->col()) && ( (col_+cols_-1) <= (xlRef->col()+xlRef->cols()-1) ) ) {
					if ( (row_ >= xlRef->row()) && ( (row_+rows_-1) <= (xlRef->row()+xlRef->rows()-1) ) ) {
					return true;
					}
				}
				return false;	
			}
		}
	}

	bool XLRef::in(Real x, Real y) const {
		if ( (x >= col_) && (x <= col_+cols_-1) ) {
			if ( (y >= row_) && (y <= row_+rows_-1) ) {
				return true;
			}
		}
		return false;
	}



}