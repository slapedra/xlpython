/*
 2008 Sebastien Lapedra
*/
#ifndef nullxlevent_hpp
#define nullxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class NullXLEventBuilder : public XLEventBuilder {
	public:
		class NullXLEvent : public XLEvent {
		public:
			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { return false; }
		};
		
		static std::string id() { return "null"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new NullXLEvent()));
				}
			}
		}


	};
	
}

#endif

