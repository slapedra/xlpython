/*
 2008 Sebastien Lapedra
*/

#ifndef xlmessageevent_hpp
#define xlmessageevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLMessageEvent : public XLCommandEvent {
	public:
		XLMessageEvent(const std::string& message, const XLSession *session)
		: xlpMessage_(message, "msg") {
			xlpMessage_.setToFree(true);
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			int result = Excel4(xlcMessage, 0, 2, TempBool(1), xlpMessage_.get());
			if (result != xlretSuccess) {
				std::cerr << "filling message error" << std::endl;
			} 
			FreeAllTempMemory();
			return false;
		}

	protected:
		XlpOper xlpMessage_;
		const XLSession *session_;
	};
	
}

#endif