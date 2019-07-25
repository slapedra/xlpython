/*
 2008 Sebastien Lapedra
*/

#ifndef xlaleertevent_hpp
#define xlalertevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLAlertEvent : public XLCommandEvent {
	public:
		XLAlertEvent(const std::string& message, const XLSession *session, bool warning = true)
		: xlpMessage_(message, "msg"), warning_(warning) {
			xlpMessage_.setToFree(true);
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			if (warning_) {
				int result = Excel4(xlcAlert, 0, 2, xlpMessage_.get(), TempNum(2));
				if (result != xlretSuccess) {
					std::cerr << "filling warning message error" << std::endl;
				} 
			} else {
				int result = Excel4(xlcAlert, 0, 2, xlpMessage_.get(), TempNum(3));
				if (result != xlretSuccess) {
					std::cerr << "filling error message error" << std::endl;
				} 
			} 
			FreeAllTempMemory();
			return !warning_;
		}

	protected:
		XlpOper xlpMessage_;
		const XLSession *session_;
		bool warning_;
	};
	
}

#endif