/*
 2008 Sebastien Lapedra
*/
#ifndef onsheetxlevent_hpp
#define onsheetxlevent_hpp

#include <xlp/xl/xlcellevents/xlevent.hpp>

namespace xlp {

	class OnSheetXLEventBuilder : public XLEventBuilder {
	public:
		class OnSheetXLEvent : public XLEvent {
		public:
			OnSheetXLEvent(const boost::shared_ptr<XLEvent>& trueevent, const boost::shared_ptr<XLEvent>& falseevent) 
			: trueevent_(trueevent), falseevent_(falseevent) {
			}

			inline bool execute(const boost::shared_ptr<XLRef>& xlCell) const { 
				if (xlCell->idSheet() == idSheet()) {
					trueevent_->execute(xlCell);
				} else {
					falseevent_->execute(xlCell);
				}
				return true;
			}

		protected:
			boost::shared_ptr<XLEvent> trueevent_;
			boost::shared_ptr<XLEvent> falseevent_;
		};

		static std::string id() { return "onsheet"; }

		inline void populate(Args& args, 
			const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& key, const std::string& nameKey, 
			const std::string& src, std::list< boost::shared_ptr<XLEvent> >& events, 
			const XLSession *session) const {
			++args.index;
			REQUIRE(args.index < args.args.size(), "not conform name");
			std::map< std::string, boost::shared_ptr<XLEventBuilder> >::const_iterator truebuilderiter = 
				args.builders->find(args.args[args.index]);
			REQUIRE(truebuilderiter != args.builders->end(), "unknow command: " << args.args[args.index]);
			std::list< boost::shared_ptr<XLEvent> > trueevents;
			truebuilderiter->second->populate(args, xlDstRef, key, nameKey, src, trueevents, session);
			REQUIRE(args.index < args.args.size(), "not conform name");
			std::map< std::string, boost::shared_ptr<XLEventBuilder> >::const_iterator falsebuilderiter = 
				args.builders->find(args.args[args.index]);
			REQUIRE(falsebuilderiter != args.builders->end(), "unknow command: " << args.args[args.index]);
			std::list< boost::shared_ptr<XLEvent> > falseevents;
			falsebuilderiter->second->populate(args, xlDstRef, key, nameKey, src, falseevents, session);
			std::list< boost::shared_ptr<XLEvent> >::iterator trueiter = trueevents.begin();
			std::list< boost::shared_ptr<XLEvent> >::iterator falseiter = falseevents.begin();
			for (Size k=0;k<xlDstRef->rows();++k) {
				for (Size j=0;j<xlDstRef->cols();++j) {
					events.push_back(boost::shared_ptr<XLEvent>(new OnSheetXLEvent(*trueiter, *falseiter)));
					++trueiter;
					++falseiter;
				}
			}
		}

	};
	
}

#endif

