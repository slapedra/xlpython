/*
 2008 Sebastien Lapedra
*/
#ifndef nopxlevent_hpp
#define nopxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class NopXLEventBuilder : public XLEventBuilder {
	public:
		class NopXLEvent : public XLEvent {
		public:
			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				return true; 
			}
			inline bool entry(const boost::shared_ptr<XLRef>& xlCell) const { 
				Excel4(xlcUndo, 0, 0);
				return true; 
			}
		};
		
		static std::string id() { return "nop"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new NopXLEvent()));
				}
			}
		}


	};
	
}

#endif

