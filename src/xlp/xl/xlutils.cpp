/*
 2008 Sebastien Lapedra
*/


#include <xlp/xl//xlutils.hpp>

namespace xlp {

	DLL_API void freeOper(XLOPER *px) {
    	if (px->xltype & xltypeStr && px->val.str)
        		delete [] px->val.str;
	    else if (px->xltype & xltypeMulti && px->val.array.lparray) {
        	int size = px->val.array.rows * px->val.array.columns;
	        for (int i=0; i<size; ++i)
        	    if (px->val.array.lparray[i].xltype & xltypeStr
	            &&  px->val.array.lparray[i].val.str)
        	        delete [] px->val.array.lparray[i].val.str;
	        delete [] px->val.array.lparray;
			} else if (px->xltype & xltypeRef && px->val.mref.lpmref) {
				delete px->val.mref.lpmref;
			}
		delete px;
	}
	
	
}



