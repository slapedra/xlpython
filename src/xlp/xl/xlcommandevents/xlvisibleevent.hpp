/*
 2008 Sebastien Lapedra
*/
#ifndef xlvisibleevent_hpp
#define xlvisibleevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLVisibleEvent : public XLCommandEvent {
	public:
		XLVisibleEvent(const std::vector<std::string>& sheets, const std::string& sheet, 
			const std::string& activesheet, bool flag)
		: sheets_(sheets), sheet_(sheet), activesheet_(activesheet), flag_(flag) {
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			for (Size i=0;i<sheets_.size();++i) {
				if (sheets_[i] != sheet_) {
					Excel4(xlcWorkbookUnhide, 0, 1, TempStrStl(sheets_[i]));
					Excel4(xlcWorkbookHide, 0, 1, TempStrStl(sheets_[i]));
				}
			}
			Excel4(xlcWorkbookUnhide, 0, 1, TempStrStl(sheet_));
			if (flag_) {
				Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(sheet_));
			} else {
				Excel4(xlcWorkbookSelect, 0, 1, TempStrStl(activesheet_));
			}
			return false;
		}

	protected:
		std::string activesheet_;
		std::string sheet_;
		std::vector<std::string> sheets_;
		bool flag_;
	};
	
}

#endif