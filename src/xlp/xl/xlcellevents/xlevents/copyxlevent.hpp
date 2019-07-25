/*
 2008 Sebastien Lapedra
*/
#ifndef copyxlevent_hpp
#define copyxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class CopyXLEventBuilder : public XLEventBuilder {
	public:
		class CopyXLEvent : public XLEvent {
		public:
			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				static XLOPER xlInt;
				xlInt.xltype = xltypeInt;
				xlInt.val.w = 6;
				Xloper xlValue;
				Excel4(xlfGetCell, &xlValue, 2, &xlInt, &(*(xlCell->oper())));
				Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(xlCell->oper())));
				Excel4(xlcFormula, 0, 2, &xlValue, &(*(xlCell->oper())));
				return true;
			}
		};
		
		static std::string id() { return "copy"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new CopyXLEvent()));
				}
			}
		}


	};
	
}

#endif

