/*
 2008 Sebastien Lapedra
*/
#ifndef xloutputevent_hpp
#define xloutputevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLOutputEvent : public XLCommandEvent {
	public:
		XLOutputEvent(const boost::shared_ptr< Range<Scalar> >& data, bool transpose,  
			const boost::shared_ptr<XLRef>& xlCell, const XLSession *session, Real time = 0.0)
		: data_(data), xlCell_(xlCell), time_(time), 
			session_(session), transpose_(transpose) {
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			if (time_ != 0) {
				while (session_->time() < time_);
			}
			XlpOper xlpData(data_, "datas", transpose_, true);
			xlpData.setToFree(true);
			int result = Excel4(xlSet, 0, 2, &(*(xlCell_->oper())), xlpData.get());
			if (result != xlretSuccess) {
				std::cerr << "xlSet error" << std::endl;
			}
			xlCell = xlCell_;
			return false;
		}

	protected:
		boost::shared_ptr< Range<Scalar> > data_;
		boost::shared_ptr<XLRef> xlCell_;
		Real time_;
		const XLSession *session_;
		bool transpose_;
	};
	
}

#endif