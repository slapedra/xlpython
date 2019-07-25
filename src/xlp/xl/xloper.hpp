/*
 2008 Sebastien Lapedra
*/

#ifndef XLP_xloper_hpp
#define XLP_xloper_hpp

#include <xlp/xl/xlcall.h>
#include <xlp/xl/framewrk.hpp>


namespace xlp {

	class DLL_API Xloper {
	public:
		Xloper(bool free = true) {
			xloper_.xltype = 0; 
			free_ = free;
		}

    ~Xloper() {
			if ( (xloper_.xltype) && free_)
				Excel(xlFree, 0, 1, &xloper_);
    }
     
		XLOPER *operator&() { return &xloper_; }
      
		const XLOPER *operator->() const { return &xloper_; }
    
		const XLOPER &operator()() const { return xloper_; }
    
	protected:
		XLOPER xloper_;
		bool free_;
	};

}

#endif

