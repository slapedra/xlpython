/*
 2008 Sebastien Lapedra
*/
#ifndef xlformulaevent_hpp
#define xlformulaevent_hpp

#include <xlp/xl/xlsession.hpp>
#include <xlp/xl/xlcommandevent.hpp>
#include <xlp/xl/xlref.hpp>

namespace xlp {
 
	class XLFormulaEvent : public XLCommandEvent {
	public:
		XLFormulaEvent(const boost::shared_ptr< Range<Scalar> >& data, bool transpose, 
			const boost::shared_ptr<XLRef>& xlCell, const XLSession *session, Real time = 0.0)
		: data_(data), transpose_(transpose), xlCell_(xlCell), time_(time), session_(session) {
		}

		inline bool execute(boost::shared_ptr<XLRef>& xlCell) const {
			if (time_ != 0) {
				while (session_->time() < time_);
			}
			XlpOper xlpData(data_, "datas", transpose_, true);
			xlpData.setToFree(true);
			boost::shared_ptr< Range<Scalar> > data2 = xlpData;
			Size row = xlCell_->row();
			Size col = xlCell_->col();
			Size minrows = std::min(xlCell_->rows(), data2->rows);
			Size mincols = std::min(xlCell_->cols(), data2->columns);
			long idSheet = xlCell_->idSheet();
			for (Size i=0;i<minrows;++i) {
				for (Size j=0;j<mincols;++j) {
					Scalar value = data2->data[i*data2->rows+j];
					boost::shared_ptr<XLRef> ref(new XLRef(row+i, col+j, 1, 1, idSheet));
					if (value.missing() || value.error()) { 
						Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(ref->oper())));
					} if (boost::apply_visitor(ScalarTypeVisitor(), value.scalar()) == ScalarString) {
						std::string str = value;
						Excel4(xlcFormula, 0, 2, TempStrStl(str), &(*(ref->oper())));
					} else {
						XlpOper xlpValue(value, "");
						Excel4(xlcFormula, 0, 2, xlpValue.get(), &(*(ref->oper())));
					}
				}
			}
			for (Size i=minrows;i<xlCell_->rows();++i) {
				for (Size j=0;j<xlCell_->cols();++j) {
					boost::shared_ptr<XLRef> ref(new XLRef(row+i, col+j, 1, 1, idSheet));
					Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(ref->oper())));
				}
			}
			for (Size i=0;i<xlCell_->rows();++i) {
				for (Size j=mincols;j<xlCell_->cols();++j) {
					boost::shared_ptr<XLRef> ref(new XLRef(row+i, col+j, 1, 1, idSheet));
					Excel4(xlcFormula, 0, 2, TempStrStl(""), &(*(ref->oper())));
				}
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
