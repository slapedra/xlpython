/*
 2008 Sebastien Lapedra
*/
#ifndef customxlcellevent_hpp
#define customxlcellevent_hpp

#include <xlp/xl/xlcellevent.hpp>
#include <xlp/xl/xlcellevents/xlevents/all.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlref.hpp>
#include <xlp/xl/xlcellevents/xlevent.hpp>
#include <boost/xpressive/xpressive_static.hpp>
#include <boost/algorithm/string.hpp> 
#include <boost/algorithm/string/trim.hpp>

namespace xlp {

	#define INSERTHANDLERBUILDER(X) \
		eventbuilders.insert(std::make_pair(X::id(), \
				boost::shared_ptr<XLEventBuilder>(new X())))
 
	class CustomXLCellEvent : public XLCellEvent {
	public:
		CustomXLCellEvent(const boost::shared_ptr<XLEvent>& entry, 
			const boost::shared_ptr<XLEvent>& dc,
			const boost::shared_ptr<XLEvent>& time,
			const boost::shared_ptr<XLEvent>& open,
			const boost::shared_ptr<XLEvent>& sheet,
			const boost::shared_ptr<XLEvent>& skey,
			const boost::shared_ptr<XLRef>& xlCell)
		: xlCell_(xlCell), entry_(entry), time_(time), 
			dc_(dc), open_(open), sheet_(sheet), skey_(skey) {
		}

		inline bool onEntry() const {
			return entry_->entry(xlCell_);
		}
		inline bool onDC() const {
			return dc_->execute(xlCell_);
		}
		inline bool onTime() const {
			return time_->execute(xlCell_);
		}
		inline bool onOpen() const {
			return open_->execute(xlCell_);
		}
		inline bool onSheet() const {
			return sheet_->execute(xlCell_);
		}
		inline bool onSKey() const {
			return skey_->execute(xlCell_);
		}


	protected:
		boost::shared_ptr<XLRef> xlCell_;
		boost::shared_ptr<XLEvent> entry_;
		boost::shared_ptr<XLEvent> dc_;
		boost::shared_ptr<XLEvent> time_;
		boost::shared_ptr<XLEvent> open_;
		boost::shared_ptr<XLEvent> sheet_;
		boost::shared_ptr<XLEvent> skey_;
	
	public:

		static bool populate(const boost::shared_ptr<XLRef>& xlDstRef, 
			const std::string& name, const std::string& nameKey, 
			const std::string& src, const boost::shared_ptr<XLCellEventMap>& cellmap, 
			const XLSession *session) {
			static boost::shared_ptr<XLEventBuilder> nullBuilder(new NullXLEventBuilder());
			static std::map<std::string, boost::shared_ptr<XLEventBuilder> > eventbuilders;
			typedef std::map<std::string, boost::shared_ptr<XLEventBuilder> >::iterator Iterator;
			static bool once = true;
			if (once) {
				INSERTHANDLERBUILDER(NullXLEventBuilder);
				INSERTHANDLERBUILDER(NopXLEventBuilder);
				INSERTHANDLERBUILDER(CopyXLEventBuilder);
				INSERTHANDLERBUILDER(TrueXLEventBuilder);
				INSERTHANDLERBUILDER(FalseXLEventBuilder);
				INSERTHANDLERBUILDER(FlagXLEventBuilder);
				INSERTHANDLERBUILDER(OneXLEventBuilder);
				INSERTHANDLERBUILDER(FlipXLEventBuilder);
				INSERTHANDLERBUILDER(DftXLEventBuilder);
				INSERTHANDLERBUILDER(TimeXLEventBuilder);
				INSERTHANDLERBUILDER(PartXLEventBuilder);
				INSERTHANDLERBUILDER(SerieXLEventBuilder);
				INSERTHANDLERBUILDER(SeekXLEventBuilder);
				once = false;
			}
			boost::shared_ptr<XLEventBuilder> entry = nullBuilder;
			boost::shared_ptr<XLEventBuilder> dc = nullBuilder;
			boost::shared_ptr<XLEventBuilder> sheet = nullBuilder;
			boost::shared_ptr<XLEventBuilder> open = nullBuilder;
			boost::shared_ptr<XLEventBuilder> time = nullBuilder;
			boost::shared_ptr<XLEventBuilder> skey = nullBuilder;
			XLEventBuilder::Args entryargs;
			XLEventBuilder::Args dcargs;
			XLEventBuilder::Args sheetargs;
			XLEventBuilder::Args openargs;
			XLEventBuilder::Args timeargs;
			XLEventBuilder::Args skeyargs;
			std::vector<std::string> tokens;
			boost::split(tokens, name, boost::is_any_of("_"));
			REQUIRE (tokens.size() > 1, "not a valid custom name: " << name);
			REQUIRE (tokens[0] == "xlp", "not a valid custom name: " << name);
			std::string key = tokens[tokens.size()-1];
			for (Size i=1;i<tokens.size()-1;++i) {
				std::vector<std::string> args;
				boost::split(args, tokens[i], boost::is_any_of("."));
				REQUIRE (args.size() > 1, "not a valid custom name: " << name);
				if (args[0] == "cell") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					entry = iter->second;
					dc = iter->second;
					skey = iter->second;
					entryargs = XLEventBuilder::Args(args, 1, &eventbuilders);
					dcargs = XLEventBuilder::Args(args, 1, &eventbuilders);
					skeyargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "scell") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					dc = iter->second;
					skey = iter->second;
					dcargs = XLEventBuilder::Args(args, 1, &eventbuilders);
					skeyargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "entry") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					entry = iter->second;
					entryargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "dc") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					dc = iter->second;
					dcargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "skey") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					skey = iter->second;
					skeyargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "time") {
					time = boost::shared_ptr<XLEventBuilder>(new OnTimeXLEventBuilder());
					timeargs = XLEventBuilder::Args(args, 0, &eventbuilders);
				} else if (args[0] == "sheet") {
					sheet = boost::shared_ptr<XLEventBuilder>(new OnSheetXLEventBuilder());
					sheetargs = XLEventBuilder::Args(args, 0, &eventbuilders);
				} else if (args[0] == "open") {
					Iterator iter = eventbuilders.find(args[1]);
					REQUIRE (iter != eventbuilders.end(), "not a valid custom name: " << name);
					open = iter->second;
					openargs = XLEventBuilder::Args(args, 1, &eventbuilders);
				} else if (args[0] == "v") {
				} else {
					FAIL ("not a valid custom name: " << name);
				}
			}
			std::list< boost::shared_ptr<XLEvent> > entryevents;
			std::list< boost::shared_ptr<XLEvent> > dcevents;
			std::list< boost::shared_ptr<XLEvent> > timeevents;
			std::list< boost::shared_ptr<XLEvent> > sheetevents;
			std::list< boost::shared_ptr<XLEvent> > openevents;
			std::list< boost::shared_ptr<XLEvent> > skeyevents;
			try {
				entry->populate(entryargs, xlDstRef, key, nameKey, src, entryevents, session);
				dc->populate(dcargs, xlDstRef, key, nameKey, src, dcevents, session);
				time->populate(timeargs, xlDstRef, key, nameKey, src, timeevents, session);
				sheet->populate(sheetargs, xlDstRef, key, nameKey, src, sheetevents, session);
				open->populate(openargs, xlDstRef, key, nameKey, src, openevents, session);
				skey->populate(skeyargs, xlDstRef, key, nameKey, src, skeyevents, session);
				std::list< boost::shared_ptr<XLEvent> >::iterator entryiter = entryevents.begin();
				std::list< boost::shared_ptr<XLEvent> >::iterator dciter = dcevents.begin();
				std::list< boost::shared_ptr<XLEvent> >::iterator timeiter = timeevents.begin();
				std::list< boost::shared_ptr<XLEvent> >::iterator sheetiter = sheetevents.begin();
				std::list< boost::shared_ptr<XLEvent> >::iterator openiter = openevents.begin();
				std::list< boost::shared_ptr<XLEvent> >::iterator skeyiter = skeyevents.begin();
				Size drow = xlDstRef->row();
				Size dcol = xlDstRef->col();
				for (Size k=0;k<xlDstRef->rows();++k) {
					for (Size j=0;j<xlDstRef->cols();++j) {
						boost::shared_ptr<XLRef> xlRef(new XLRef(drow+k, dcol+j, 1, 1, xlDstRef->idSheet()));
						boost::shared_ptr<XLCellEvent> handler(new CustomXLCellEvent(
							*entryiter, *dciter, *timeiter, *openiter, *sheetiter, *skeyiter, xlRef));
						cellmap->add(drow+k, dcol+j, xlDstRef->idSheet(), handler);
						++entryiter;
						++dciter;
						++timeiter;
						++sheetiter;
						++openiter;
						++skeyiter;
					}
				}
			} catch(...) {
				return false;
			}
			return true;
		}


	};
	
}

#endif
