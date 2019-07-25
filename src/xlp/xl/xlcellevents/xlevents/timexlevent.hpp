/*
 2008 Sebastien Lapedra
*/
#ifndef timexlevent_hpp
#define timexlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class TimeXLEventBuilder : public XLEventBuilder {
	public:
		class TimeXLEvent : public XLEvent {
		public:
			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				Excel4(xlcFormula, 0, 2, &time(), &(*(xlCell->oper())));
				return true;
			}
		};
		
		static std::string id() { return "time"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new TimeXLEvent()));
				}
			}
		}
	};
	
}

#endif

