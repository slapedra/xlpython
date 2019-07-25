/*
 2008 Sebastien Lapedra
*/

#ifndef xlnameevent_hpp
#define xlnameevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLNameEvent : public XLCommandEvent {
	public:
		XLNameEvent(const std::string& name, const boost::shared_ptr<XLRef>& xlCell, const XLSession *session)
		: xlCell_(xlCell), xlpName_(name, "name") {
			xlpName_.setToFree(true);
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			int result = Excel4(xlcDefineName, 0, 2, xlpName_.get(), &(*(xlCell_->oper())));
			if (result != xlretSuccess) {
				std::cerr << "xlcDefineName" << std::endl;
			} 
			return false;
		}

	protected:
		XlpOper xlpName_;
		boost::shared_ptr<XLRef> xlCell_;
		const XLSession *session_;
	};
	
}

#endif