/*
 2008 Sebastien Lapedra
*/
#ifndef xlselectevent_hpp
#define xlselectevent_hpp

#include <xlp/xl/xlref.hpp>
#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>


namespace xlp {
 
	class XLSelectEvent : public XLCommandEvent {
	public:
		XLSelectEvent(const boost::shared_ptr<XLRef>& xlCell, const XLSession *session)
		: session_(session), xlCell_(xlCell) { 
		}

		inline bool execute() const {
			Excel4(xlcSelect, 0, 1, &(*(xlCell_->oper())));
			session_->unprotect(false);
			return false;
		}

	protected:
		boost::shared_ptr<XLRef> xlCell_;
		const XLSession *session_;
	};
	
}

#endif
